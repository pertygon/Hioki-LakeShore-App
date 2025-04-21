import sys, time, pyvisa as visa
import numpy as np, pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTextEdit, QProgressBar,
    QComboBox, QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt
from instrument import (
    Lakeshore335, Hioki3536,
    MockLakeshore335, MockHioki3536
)
from utils import save_results

class SweepApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Impedance Sweep")
        self.resize(550, 450)

        self.paused = False
        self.stopped = False
        self.freqs = []
        self.temps = []

        self._build_ui()
        self._detect_devices()

    def _detect_devices(self):
        rm = visa.ResourceManager()
        resources = rm.list_resources()
        lake_list = ["Mock"]
        hioki_list = ["Mock"]

        for res in resources:
            try:
                dev = rm.open_resource(res)
                dev.write_termination = '\n'
                dev.read_termination  = '\n'
                idn = dev.query("*IDN?")
                dev.close()
                if "Lakeshore" in idn or "335" in idn:
                    lake_list.append(res)
                elif "Hioki" in idn or "3536" in idn:
                    hioki_list.append(res)
            except:
                continue

        self.cb_lake.clear();  self.cb_lake.addItems(lake_list)
        self.cb_hioki.clear(); self.cb_hioki.addItems(hioki_list)

    def _build_ui(self):
        layout = QVBoxLayout()
        # — wybór urządzeń
        dev_layout = QHBoxLayout()
        dev_layout.addWidget(QLabel("Lakeshore:"))
        self.cb_lake = QComboBox(); dev_layout.addWidget(self.cb_lake)
        dev_layout.addWidget(QLabel("Hioki:"))
        self.cb_hioki = QComboBox(); dev_layout.addWidget(self.cb_hioki)
        layout.addLayout(dev_layout)

        # — przycisk ładowania Excel (2 kolumny: freq|temp)
        xl_layout = QHBoxLayout()
        self.btn_load = QPushButton("Wczytaj Excel (Freq | Temp)")
        self.btn_load.clicked.connect(self._load_ranges)
        xl_layout.addWidget(self.btn_load)
        self.lbl_ranges = QLabel("Brak danych")
        xl_layout.addWidget(self.lbl_ranges)
        layout.addLayout(xl_layout)

        # — Start / Pause / Stop
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

        # — pasek postępu
        self.progress = QProgressBar()
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        # — log
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        self.setLayout(layout)

    def _load_ranges(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Wczytaj Excel", "", "Excel (*.xlsx *.xls)"
        )
        if not path:
            return
        df = pd.read_excel(path)
        # bierzemy obie kolumny jako niezależne wektory
        self.freqs = df.iloc[:,0].dropna().tolist()
        self.temps = df.iloc[:,1].dropna().tolist()
        self.lbl_ranges.setText(
            f"{len(self.temps)} temperature i {len(self.freqs)} częstotliwości"
        )

    def _toggle_pause(self):
        self.paused = not self.paused
        self.btn_pause.setText("Resume" if self.paused else "Pause")

    def _stop(self):
        self.stopped = True
        self.log.append("**STOP** – kończę pętlę.")
        self.progress.setValue(0)    # reset progress

    def run_sweep(self):
        if not self.freqs or not self.temps:
            QMessageBox.warning(self, "Błąd", "Wczytaj Excel z freq i temp")
            return

        # blokada/odblokowanie przycisków
        self.btn_start.setEnabled(False)
        self.btn_pause .setEnabled(True)
        self.btn_stop  .setEnabled(True)
        self.stopped = False
        self.paused  = False
        self.progress.setValue(0)

        # wykryte urządzenia
        lake_sel = self.cb_lake.currentText()
        hioki_sel = self.cb_hioki.currentText()
        self.lake  = MockLakeshore335() if lake_sel=="Mock"  else Lakeshore335(lake_sel)
        self.hioki = MockHioki3536()  if hioki_sel=="Mock" else Hioki3536(hioki_sel)

        total = len(self.temps) * len(self.freqs)
        step = 0

        for T in self.temps:
            if self.stopped: break

            self.log.append(f"> Ustawiam T = {T:.2f} K")
            self.lake.set_temperature(T)
            # czekaj na ±0.5 K
            while abs(self.lake.get_temperature() - T) > 0.5:
                if self.stopped: break
                while self.paused:
                    time.sleep(0.1)
                    QApplication.processEvents()
                time.sleep(1)

            if self.stopped: break

            measurements = []
            for f in self.freqs:
                if self.stopped: break
                while self.paused:
                    time.sleep(0.1)
                    QApplication.processEvents()

                self.log.append(f"  f = {f:.1f} Hz …")
                self.hioki.set_frequency(f)
                time.sleep(0.2)
                data = self.hioki.measure_all()

                step += 1
                entry = {
                    'Lp.':   step,
                    'Freq':  f,
                    'Temp':  T,
                    'PHASe': data['Phase'],
                    'Cp':    data['Cp'],
                    'D':     data['D'],
                    'Rp':    data['Rp'],
                }
                measurements.append(entry)

                self.progress.setValue(int(step/total*100))
                QApplication.processEvents()

            # zapisz dane dla tej temperatury
            df_T = pd.DataFrame(measurements)
            filename = f"{T:.0f}.csv"
            save_results(df_T, filename)
            self.log.append(f"Zapisano: {filename}")

        # po skończeniu (bez Stop) reset paska
        if not self.stopped:
            self.progress.setValue(0)

        # przywróć stan przycisków
        self.btn_start.setEnabled(True)
        self.btn_pause .setEnabled(False)
        self.btn_stop  .setEnabled(False)
        self.log.append("=== Pomiary zakończone ===")
