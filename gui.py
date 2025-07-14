# gui.py

import sys, os
import pyvisa
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QProgressBar, QComboBox, QListWidget,
    QListWidgetItem, QFileDialog, QMessageBox, QSpinBox,
    QDoubleSpinBox, QAbstractItemView, QSplashScreen
)
from PyQt5.QtCore import Qt, QThread
from PyQt5.QtGui import QPixmap
from instrument import Lakeshore335, Hioki3536, MockLakeshore335, MockHioki3536
from measurement import SweepWorker

class SweepApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Impedance Sweep")
        self.resize(600, 580)

        # stan pomiaru
        self.lake       = None
        self.output_dir = None
        self.worker     = None
        self.thread     = None
        self.measuring  = False

        self._build_ui()
        self._detect_devices()
        self.cb_lake.currentIndexChanged.connect(self._init_lake)
        self._init_lake()

    def _build_ui(self):
        layout = QVBoxLayout()

        # urządzenia
        dev_layout = QHBoxLayout()
        dev_layout.addWidget(QLabel("Lakeshore:"))
        self.cb_lake = QComboBox(); dev_layout.addWidget(self.cb_lake)
        dev_layout.addWidget(QLabel("Hioki:"))
        self.lst_hioki = QListWidget()
        self.lst_hioki.setSelectionMode(QAbstractItemView.MultiSelection)
        dev_layout.addWidget(self.lst_hioki)

        # symulatory
        sim_layout = QVBoxLayout()
        self.btn_add_mock = QPushButton("Dodaj symulowany Hioki")
        self.btn_add_mock.clicked.connect(self._add_mock)
        sim_layout.addWidget(self.btn_add_mock)
        self.btn_remove_mock = QPushButton("Usuń symulowany Hioki")
        self.btn_remove_mock.clicked.connect(self._remove_mock)
        sim_layout.addWidget(self.btn_remove_mock)
        dev_layout.addLayout(sim_layout)

        layout.addLayout(dev_layout)

        # zakresy i parametry
        params = QHBoxLayout()
        self.btn_load = QPushButton("Wczytaj Excel\n(Temp|Freq)")
        self.btn_load.clicked.connect(self._load_ranges)
        params.addWidget(self.btn_load)
        self.lbl_ranges = QLabel("Brak danych"); params.addWidget(self.lbl_ranges)

        self.btn_choose_dir = QPushButton("Wybierz folder wyników")
        self.btn_choose_dir.clicked.connect(self._choose_folder)
        params.addWidget(self.btn_choose_dir)
        self.lbl_folder = QLabel("Brak folderu"); params.addWidget(self.lbl_folder)

        params.addWidget(QLabel("Czas stab. [s]:"))
        self.sb_stab = QSpinBox(); self.sb_stab.setRange(1,3600); self.sb_stab.setValue(30)
        params.addWidget(self.sb_stab)

        params.addWidget(QLabel("Tol [K]:"))
        self.ds_tol = QDoubleSpinBox(); self.ds_tol.setRange(0.01,5)
        self.ds_tol.setSingleStep(0.01); self.ds_tol.setValue(0.1)
        params.addWidget(self.ds_tol)

        params.addWidget(QLabel("Offset [K]:"))
        self.ds_off = QDoubleSpinBox(); self.ds_off.setRange(-5,5)
        self.ds_off.setSingleStep(0.01); self.ds_off.setValue(0.0)
        params.addWidget(self.ds_off)

        layout.addLayout(params)

        # heater
        heater = QHBoxLayout()
        self.btn_heat_on  = QPushButton("Heater ON");  self.btn_heat_off = QPushButton("Heater OFF")
        self.btn_heat_on.setEnabled(False); self.btn_heat_off.setEnabled(False)
        self.btn_heat_on.clicked.connect(self._heater_on)
        self.btn_heat_off.clicked.connect(self._heater_off)
        heater.addWidget(self.btn_heat_on); heater.addWidget(self.btn_heat_off)
        layout.addLayout(heater)

        # sterowanie
        ctrl = QHBoxLayout()
        self.btn_start = QPushButton("Start")
        self.btn_pause = QPushButton("Pause"); self.btn_pause.setCheckable(True); self.btn_pause.setEnabled(False)
        self.btn_stop  = QPushButton("Stop");   self.btn_stop .setEnabled(False)
        ctrl.addWidget(self.btn_start); ctrl.addWidget(self.btn_pause); ctrl.addWidget(self.btn_stop)
        layout.addLayout(ctrl)

        self.btn_start.clicked.connect(self.run_sweep)
        self.btn_pause.toggled.connect(self._on_pause_toggled)
        self.btn_pause.toggled.connect(lambda p: self.worker.pause(p) if self.worker else None)
        self.btn_stop.clicked.connect(lambda: (self.worker.stop() if self.worker else None,
                                              self.lbl_status.setText("Zakończono pomiary"),
                                              self.progress.setValue(0)))

        # pasek + status
        self.progress   = QProgressBar()
        self.lbl_status = QLabel("Gotowy"); self.lbl_status.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.progress); layout.addWidget(self.lbl_status)

        self.setLayout(layout)

    def _detect_devices(self):
        # splash z kursorem "busy"
        splash_pix = QPixmap(300,100); splash_pix.fill(Qt.white)
        splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
        splash.showMessage("Wyszukuję urządzenia…", Qt.AlignCenter, Qt.black)
        splash.show()
        QApplication.processEvents()
        QApplication.setOverrideCursor(Qt.WaitCursor)

        rm = pyvisa.ResourceManager()
        resources = rm.list_resources()
        lakes, hiokis = ["Mock"], []
        for r in resources:
            try:
                inst = rm.open_resource(r)
                inst.baud_rate=57600; inst.data_bits=7
                inst.stop_bits=pyvisa.constants.StopBits.one; inst.parity=pyvisa.constants.Parity.odd
                inst.timeout=2000; inst.write_termination='\r\n'; inst.read_termination='\r\n'
                if "335" in inst.query("*IDN?").upper():
                    lakes.append(r); inst.close(); continue
                inst.close()
            except: pass
            try:
                inst = rm.open_resource(r)
                inst.baud_rate=19200; inst.data_bits=8
                inst.stop_bits=pyvisa.constants.StopBits.one; inst.parity=pyvisa.constants.Parity.none
                inst.timeout=2000; inst.write_termination='\r\n'; inst.read_termination='\r\n'
                if "3536" in inst.query("*IDN?").upper():
                    hiokis.append(r)
                inst.close()
            except: pass

        QApplication.restoreOverrideCursor()
        splash.finish(self)

        self.cb_lake.clear(); self.cb_lake.addItems(lakes)
        self.lst_hioki.clear()
        for h in hiokis:
            item = QListWidgetItem(h); item.setData(Qt.UserRole,h); item.setSelected(True)
            self.lst_hioki.addItem(item)

    def _init_lake(self):
        if hasattr(self, 'lake') and self.lake:
            try: self.lake.close()
            except: pass
        res = self.cb_lake.currentText()
        if res.startswith("Mock"):
            self.lake,ok = MockLakeshore335(),False
        else:
            self.lake,ok = Lakeshore335(res),True
        self.btn_heat_on .setEnabled(ok)
        self.btn_heat_off.setEnabled(ok)

    def _add_mock(self):
        item = QListWidgetItem("Mock"); item.setData(Qt.UserRole,"Mock"); item.setSelected(True)
        self.lst_hioki.addItem(item)

    def _remove_mock(self):
        for it in self.lst_hioki.selectedItems():
            if it.data(Qt.UserRole)=="Mock":
                self.lst_hioki.takeItem(self.lst_hioki.row(it))

    def _load_ranges(self):
        path,_ = QFileDialog.getOpenFileName(self,"Wczytaj Excel","", "Excel (*.xlsx *.xls)")
        if not path: return
        df = pd.read_excel(path, header=None)
        self.temps = df.iloc[:,0].dropna().tolist()
        self.freqs = df.iloc[:,1].dropna().tolist()
        self.lbl_ranges.setText(f"{len(self.temps)} temp, {len(self.freqs)} freq")

    def _choose_folder(self):
        path = QFileDialog.getExistingDirectory(self,"Folder wyników")
        if path: self.output_dir,path = path,path

    def _heater_on(self):
        try:
            self.lake.enable_heater(); self.lbl_status.setText("Heater ON")
        except Exception as e:
            QMessageBox.warning(self,"Błąd",f"Heater ON:\n{e}")

    def _heater_off(self):
        try:
            self.lake.disable_heater(); self.lbl_status.setText("Heater OFF")
        except Exception as e:
            QMessageBox.warning(self,"Błąd",f"Heater OFF:\n{e}")

    def run_sweep(self):
        if self.measuring: return
        self.measuring = True
        self.lbl_status.setText("Start pomiarów…")

        if not getattr(self,'output_dir',None):
            fld = QFileDialog.getExistingDirectory(self,"Wybierz folder wyników", os.getcwd())
            if not fld:
                QMessageBox.warning(self,"Błąd","Nie wybrano folderu – anulowano.")
                self.lbl_status.setText("Anulowano"); self.measuring=False; return
            self.output_dir = fld

        if not self.lake:
            QMessageBox.warning(self,"Błąd","Wybierz Lakeshore!"); self.lbl_status.setText("Gotowy"); self.measuring=False; return
        hioki = [it.text() for it in self.lst_hioki.selectedItems()]
        if self.cb_lake.currentText()=="Mock" and not hioki:
            QMessageBox.warning(self,"Błąd","Wybierz ≥1 Hioki lub Mock!")
            self.lbl_status.setText("Gotowy"); self.measuring=False; return
        if not getattr(self,'temps',None) or not self.freqs:
            QMessageBox.warning(self,"Błąd","Wczytaj Excel!")
            self.lbl_status.setText("Gotowy"); self.measuring=False; return

        self.btn_start.setEnabled(False)
        self.btn_pause.setEnabled(True)
        self.btn_stop .setEnabled(True)
        self.btn_pause.setChecked(False)
        self.btn_pause.setText("Pause")

        hioki_objs = [(n, MockHioki3536() if n=="Mock" else Hioki3536(n)) for n in hioki]
        self.worker = SweepWorker(
            lake=self.lake, hiokis=hioki_objs,
            temps=self.temps, freqs=self.freqs,
            stabilize_time=self.sb_stab.value(),
            tol=self.ds_tol.value(),
            offset=self.ds_off.value(),
            output_dir=self.output_dir
        )
        self.thread = QThread(self)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.status.connect(self.lbl_status.setText)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self._on_finished)

        self.thread.start()

    def _on_pause_toggled(self, paused: bool):
        if self.worker:
            self.worker.pause(paused)
        self.btn_pause.setText("Resume" if paused else "Pause")
        self.lbl_status.setText("Zatrzymano wykonywanie pomiarów" if paused else "Wznawianie pomiarów...")

    def _on_stop_clicked(self):
        if self.worker:
            self.worker.stop()
        self.lbl_status.setText("Zakończono pomiary")
        self.progress.setValue(0)

    def _on_finished(self):
        self.measuring = False
        self.thread     = None
        self.worker     = None
        self.lbl_status.setText("Pomiary zakończone")
        self.progress.setValue(0)
        self.btn_start.setEnabled(True)
        self.btn_pause.setEnabled(False)
        self.btn_pause.setChecked(False)
        self.btn_pause.setText("Pause")
        self.btn_stop .setEnabled(False)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # splash-screen z kółkiem kursora
    pix = QPixmap(300,100)
    pix.fill(Qt.white)
    splash = QSplashScreen(pix, Qt.WindowStaysOnTopHint)
    splash.showMessage("Wyszukuję urządzenia…", Qt.AlignCenter, Qt.black)
    splash.show()
    app.processEvents()

    window = SweepApp()
    splash.finish(window)

    window.show()
    sys.exit(app.exec_())
