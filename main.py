# main.py

import sys
import os
import time
import pyvisa as visa
import numpy as np
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTextEdit, QProgressBar,
    QComboBox, QListWidget, QListWidgetItem, QAbstractItemView,
    QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt
from instrument import Lakeshore335, Hioki3536, MockLakeshore335, MockHioki3536
from utils import save_results

class SweepApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Impedance Sweep")
        self.resize(600, 480)

        self.paused = False
        self.stopped = False
        self.freqs = []
        self.temps = []
        self.mock_count = 0

        self._build_ui()
        self._detect_devices()

    def _build_ui(self):
        layout = QVBoxLayout()

        # — wybór urządzeń
        dev_layout = QHBoxLayout()
        dev_layout.addWidget(QLabel("Lakeshore:"))
        self.cb_lake = QComboBox()
        dev_layout.addWidget(self.cb_lake)

        dev_layout.addWidget(QLabel("Hioki:"))
        self.lst_hioki = QListWidget()
        self.lst_hioki.setSelectionMode(QAbstractItemView.MultiSelection)
        dev_layout.addWidget(self.lst_hioki)

        # — przyciski do simulacji
        sim_btns = QVBoxLayout()
        self.btn_add_mock = QPushButton("Dodaj symulowany miernik")
        self.btn_add_mock.clicked.connect(self._add_mock)
        sim_btns.addWidget(self.btn_add_mock)
        self.btn_remove_mock = QPushButton("Usuń symulowany miernik")
        self.btn_remove_mock.clicked.connect(self._remove_mock)
        sim_btns.addWidget(self.btn_remove_mock)
        dev_layout.addLayout(sim_btns)

        layout.addLayout(dev_layout)

        # — wczytywanie zakresów
        xl_layout = QHBoxLayout()
        self.btn_load = QPushButton("Wczytaj Excel (Temp | Freq)")
        self.btn_load.clicked.connect(self._load_ranges)
        xl_layout.addWidget(self.btn_load)
        self.lbl_ranges = QLabel("Brak danych")
        xl_layout.addWidget(self.lbl_ranges)
        layout.addLayout(xl_layout)

        # — przyciski start/pause/stop
        btn_layout = QHBoxLayout()
        self.btn_start = QPushButton("Start")
        self.btn_start.clicked.connect(self.run_sweep)
        btn_layout.addWidget(self.btn_start)
        self.btn_pause = QPushButton("Pause")
        self.btn_pause.setEnabled(False)
        self.btn_pause.clicked.connect(self._toggle_pause)
        btn_layout.addWidget(self.btn_pause)
        self.btn_stop = QPushButton("Stop")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self._stop)
        btn_layout.addWidget(self.btn_stop)
        layout.addLayout(btn_layout)

        # — pasek postępu i log
        self.progress = QProgressBar()
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        self.setLayout(layout)

    def _add_mock(self):
        """Dodaje unikalny symulowany miernik do listy."""
        self.mock_count += 1
        name = f"Mock{self.mock_count}"
        item = QListWidgetItem(name)
        item.setData(Qt.UserRole, "Mock")
        self.lst_hioki.addItem(item)

    def _remove_mock(self):
        """Usuwa zaznaczone symulowane mierniki."""
        for item in list(self.lst_hioki.selectedItems()):
            if item.data(Qt.UserRole) == "Mock":
                self.lst_hioki.takeItem(self.lst_hioki.row(item))

    def _detect_devices(self):
        rm = visa.ResourceManager()
        resources = rm.list_resources()
        lake_list = ["Mock"]
        hioki_list = []

        for res in resources:
            if res.upper().startswith("ASRL"):
                # Lakeshore 335 @57600,7E1
                try:
                    inst = rm.open_resource(res)
                    inst.baud_rate = 57600
                    inst.data_bits = 7
                    inst.stop_bits = visa.constants.StopBits.one
                    inst.parity = visa.constants.Parity.odd
                    inst.timeout = 10000
                    inst.write_termination = '\r\n'
                    inst.read_termination  = '\r\n'
                    resp = inst.query("*IDN?").strip().upper()
                    inst.close()
                    if any(tag in resp for tag in ("LSCI","335","LAKESHORE")):
                        lake_list.append(res)
                        continue
                except:
                    pass
                # Hioki 3536 @19200,8N1
                try:
                    inst = rm.open_resource(res)
                    inst.baud_rate = 19200
                    inst.data_bits = 8
                    inst.stop_bits = visa.constants.StopBits.one
                    inst.parity = visa.constants.Parity.none
                    inst.timeout = 2000
                    inst.write_termination = '\r\n'
                    inst.read_termination  = '\r\n'
                    resp = inst.query("*IDN?").strip().upper()
                    inst.close()
                    if any(tag in resp for tag in ("3536","HIOKI")):
                        hioki_list.append(res)
                        continue
                except:
                    pass
            else:
                try:
                    inst = rm.open_resource(res)
                    inst.timeout = 2000
                    inst.write_termination = '\r\n'
                    inst.read_termination  = '\r\n'
                    resp = inst.query("*IDN?").strip().upper()
                    inst.close()
                    if any(tag in resp for tag in ("LSCI","335","LAKESHORE")):
                        lake_list.append(res)
                    elif any(tag in resp for tag in ("3536","HIOKI")):
                        hioki_list.append(res)
                except:
                    pass

        self.cb_lake.clear()
        self.cb_lake.addItems(lake_list)
        self.lst_hioki.clear()
        for h in hioki_list:
            item = QListWidgetItem(h)
            item.setData(Qt.UserRole, h)
            self.lst_hioki.addItem(item)

    def _load_ranges(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Wczytaj Excel", "", "Excel (*.xlsx *.xls)"
        )
        if not path:
            return
        df = pd.read_excel(path)
        self.temps = df.iloc[:,0].dropna().tolist()
        self.freqs = df.iloc[:,1].dropna().tolist()
        self.lbl_ranges.setText(f"{len(self.temps)} temperatur, {len(self.freqs)} częstotliwości")

    def _toggle_pause(self):
        self.paused = not self.paused
        self.btn_pause.setText("Resume" if self.paused else "Pause")

    def _stop(self):
        self.stopped = True
        self.log.append("**STOP** – kończę pętlę.")
        self.progress.setValue(0)

    def run_sweep(self):
        lake_res = self.cb_lake.currentText()
        hioki_items = self.lst_hioki.selectedItems()
        if lake_res=="Mock" and not hioki_items:
            QMessageBox.warning(self, "Błąd", "Wybierz ≥1 miernik Hioki lub Mock.")
            return
        if not self.temps or not self.freqs:
            QMessageBox.warning(self, "Błąd", "Wczytaj Excel z zakresami.")
            return

        # blokada UI
        self.btn_start.setEnabled(False)
        self.btn_pause .setEnabled(True)
        self.btn_stop  .setEnabled(True)
        self.paused = False
        self.stopped = False
        self.progress.setValue(0)

        # inicjalizacja Lakeshore
        self.lake = MockLakeshore335() if lake_res.startswith("Mock") else Lakeshore335(lake_res)

        # inicjalizacja Hioki
        hiokis = []
        for item in hioki_items:
            name = item.text()
            role = item.data(Qt.UserRole)
            inst = MockHioki3536() if role=="Mock" else Hioki3536(name)
            hiokis.append((name, inst))

        total = len(self.temps)*len(self.freqs)*len(hiokis)
        step = 0

        for T in self.temps:
            if self.stopped: break
            self.log.append(f"> Ustawiam T = {T:.2f} K")
            self.lake.set_temperature(T)
            while abs(self.lake.get_temperature() - T) > 0.5:
                if self.stopped: break
                while self.paused:
                    time.sleep(0.1)
                    QApplication.processEvents()
                time.sleep(1)
            if self.stopped: break

            # przygotuj listy wyników per miernik
            data_per_meter = {name: [] for name,_ in hiokis}

            for f in self.freqs:
                if self.stopped: break
                while self.paused:
                    time.sleep(0.1)
                    QApplication.processEvents()

                for name, meter in hiokis:
                    self.log.append(f"  [{name}] f = {f:.1f} Hz …")
                    meter.set_frequency(f)
                    time.sleep(0.2)
                    meas = meter.measure_all()
                    step += 1
                    entry = {
                        'Lp.': step,
                        'Freq': f,
                        'Temp': T,
                        'PHASe': meas['Phase'],
                        'Cp': meas['Cp'],
                        'D': meas['D'],
                        'Rp': meas['Rp'],
                    }
                    data_per_meter[name].append(entry)
                    self.progress.setValue(int(step/total*100))
                    QApplication.processEvents()

            # zapis wyników: osobny folder dla każdego miernika
            for name,_ in hiokis:
                folder = os.path.join("results", name.replace("::","_"))
                os.makedirs(folder, exist_ok=True)
                df_res = pd.DataFrame(data_per_meter[name])
                filename = f"T{T:.1f}K.csv"
                save_results(df_res, filename, out_folder=folder)
                self.log.append(f"✅ {name}: zapisano {folder}/{filename}")

        # przywrócenie UI
        if not self.stopped:
            self.progress.setValue(0)
            self.log.append("=== Pomiary zakończone ===")
        self.btn_start.setEnabled(True)
        self.btn_pause .setEnabled(False)
        self.btn_stop  .setEnabled(False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = SweepApp()
    w.show()
    sys.exit(app.exec_())
