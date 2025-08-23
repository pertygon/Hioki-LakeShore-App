# measurement.py

import os, time
import pandas as pd
from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot

class SweepWorker(QObject):
    status   = pyqtSignal(str)
    progress = pyqtSignal(int)
    finished = pyqtSignal()

    def __init__(self, lake, hiokis, temps, freqs,
                 stabilize_time, tol, offset, output_dir):
        super().__init__()
        self.lake       = lake
        self.hiokis     = hiokis
        self.temps      = temps
        self.freqs      = freqs
        self.stab       = stabilize_time
        self.tol        = tol
        self.offset     = offset
        self.output_dir = output_dir
        self._stopped = False
        self._paused  = False

    @pyqtSlot()
    def stop(self):
        self._stopped = True

    @pyqtSlot(bool)
    def pause(self, paused: bool):
        self._paused = paused

    def run(self):
        total = len(self.temps) * len(self.freqs) * len(self.hiokis)
        step  = 0

        #włącz grzałkę
        try: self.lake.enable_heater()
        except: pass

        for T in self.temps:
            if self._stopped:
                break

            # ustawienie punktu i informacja
            self.status.emit(f"Setpoint {T:.2f} K")
            self.lake.set_temperature(T)

            # przygotuj dane tylko dla tej temperatury
            data_per = {name: [] for name,_ in self.hiokis}

            # czekaj aż temperatura ustabilizuje się przez self.stab sekund
            first = True
            last_within = None
            while not self._stopped:
                if self._paused:
                    continue

                raw = self.lake.get_temperature()
                curr = raw + self.offset
                self.status.emit(f"T={curr:.2f} K")

                now = time.time()
                if first:
                    # jeśli już w tol, traktuj, że "było" self.stab sekund temu
                    last_within = now if abs(curr - T) > self.tol else now - self.stab
                    first = False

                if abs(curr - T) > self.tol:
                    last_within = now

                if now - last_within >= self.stab:
                    break


            if self._stopped:
                break

            # pomiary Hioki przy każdej częstotliwości
            for f in self.freqs:
                if self._stopped:
                    break
                while self._paused:
                    pass

                for name, meter in self.hiokis:
                    self.status.emit(f"[{name}] f={f:.1f} Hz")
                    meter.set_frequency(f)
                    meas = meter.measure_all()

                    step += 1
                    entry = {
                        'Lp.':   step,
                        'Freq':  f,
                        'Temp':  T,
                        'PHASe': meas['Phase'],
                        'Cp':    meas['Cp'],
                        'D':     meas['D'],
                        'Rp':    meas['Rp'],
                    }
                    data_per[name].append(entry)
                    self.progress.emit(int(step/total*100))

            # zapis pliku dla tej temperatury i każdego miernika
            for name,_ in self.hiokis:
                folder = os.path.join(self.output_dir, name.replace("::", "_"))
                os.makedirs(folder, exist_ok=True)
                df = pd.DataFrame(data_per[name])
                filename = f"{T}.csv"
                df.to_csv(os.path.join(folder, filename), index=False)

        # po wszystkim (lub stop) – schłódź i wyłącz grzałkę
        if not self._stopped:
            try:
                self.status.emit("Cooldown → 300 K")
                self.lake.set_temperature(300.0)
                self.lake.disable_heater()
            except:
                pass
        self.lake.set_temperature(300.0)       
        self.finished.emit()
    def manual_measure(self):
        """
        Ręczny pomiar
        """
        try:
            curr_temp = float(self.lake.get_temperature()) + self.offset
        except Exception as e:
            raise RuntimeError(f"Nie udało się odczytać temperatury: {e}")

        results = {}
        for name, meter in self.hiokis:
            data = []
            for f in self.freqs:
                try:
                    meter.set_frequency(f)
                    meas = meter.measure_all()
                    entry = {
                        "Freq":  f,
                        "Temp":  curr_temp,
                        "Phase": meas["Phase"],
                        "Cp":    meas["Cp"],
                        "D":     meas["D"],
                        "Rp":    meas["Rp"],
                    }
                    data.append(entry)
                except Exception as e:
                    print(f"[WARN] Pomiar {name} f={f} Hz nie powiódł się: {e}")
            results[name] = data

        # zapis CSV – osobny plik dla każdego Hioki
        for name, rows in results.items():
            folder = os.path.join(self.output_dir, name.replace("::", "_"))
            os.makedirs(folder, exist_ok=True)
            filename = os.path.join(folder, f"Ręczny_{curr_temp:.2f}.csv")
            pd.DataFrame(rows).to_csv(filename, index=False)

        return curr_temp, results
