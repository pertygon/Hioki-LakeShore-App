# instrument.py

import pyvisa
import time
import random
from pyvisa import constants

class MockLakeshore335:
    """Symulator kontrolera temperatury."""
    def __init__(self, *args, **kwargs):
        self._target = 300.0

    def set_temperature(self, T):
        print(f"[MOCK Lake] ustawiam temperaturę na {T:.2f} K")
        self._target = T

    def get_temperature(self):
        current = self._target + random.uniform(-0.1, 0.1)
        print(f"[MOCK Lake] odczyt temperatury: {current:.2f} K")
        return current

class MockHioki3536:
    """Symulator miernika impedancji."""
    def __init__(self, *args, **kwargs):
        pass

    def set_frequency(self, freq_hz):
        print(f"[MOCK Hioki] ustawiam częstotliwość na {freq_hz:.1f} Hz")

    def measure_all(self):
        data = {
            'Phase': random.uniform(-180, 180),
            'Cp':    random.uniform(1e-12, 1e-6),
            'D':     random.uniform(0, 1),
            'Rp':    random.uniform(1e2, 1e6),
        }
        print(f"[MOCK Hioki] pomiar: {data}")
        return data

class Lakeshore335:
    def __init__(self, resource_name):
        rm = pyvisa.ResourceManager()
        if resource_name.upper().startswith("ASRL"):
            # USB→COM emuluje RS-232C @57600 baud, 7 data bits, odd parity
            self.dev = rm.open_resource(
                resource_name,
                baud_rate=57600,
                data_bits=7,
                stop_bits=constants.StopBits.one,
                parity=constants.Parity.odd
            )
        else:
            self.dev = rm.open_resource(resource_name)
        self.dev.timeout = 2000
        self.dev.write_termination = '\r\n'
        self.dev.read_termination  = '\r\n'

    def set_temperature(self, T):
        self.dev.write(f"SETP 1,{T:.2f}")

    def get_temperature(self, retries=5, delay=0.2):
        last = ""
        for _ in range(retries):
            resp = self.dev.query("KRDG? 1").strip()
            last = resp
            if resp:
                try:
                    return float(resp)
                except ValueError:
                    for part in resp.replace('+','').split(','):
                        try:
                            return float(part)
                        except ValueError:
                            continue
            time.sleep(delay)
        raise RuntimeError(f"Nie odczytano temperatury (ostatnia: {last!r})")

class Hioki3536:
    def __init__(self, resource_name):
        rm = pyvisa.ResourceManager()
        if resource_name.upper().startswith("ASRL"):
            self.dev = rm.open_resource(
                resource_name,
                baud_rate=19200,
                data_bits=8,
                stop_bits=constants.StopBits.one,
                parity=constants.Parity.none
            )
        else:
            self.dev = rm.open_resource(resource_name)
        self.dev.timeout = 2000
        self.dev.write_termination = '\r\n'
        self.dev.read_termination  = '\r\n'

        # ustaw tryb CPD, ASCII, natychmiastowy trigger
        for cmd in ("FUNC 'CPD'", "TRIG:SOUR IMM", "FORM:DATA ASCII"):
            try:
                self.dev.write(cmd)
            except:
                pass

    def set_frequency(self, freq_hz):
        self.dev.write(f"FREQ {freq_hz:.0f}")

    def measure_all(self):
        """
        Wyzwala pomiar i pobiera ostatni zestaw Phase,Cp,D,Rp
        z odpowiedzi MEASure? ALL (panele rozdzielone '/').
        """
        # 1) wyzwól pomiar
        try:
            self.dev.write("*TRG")
        except:
            pass

        # 2) poczekaj na zakończenie
        try:
            self.dev.query("*OPC?")
        except:
            time.sleep(0.5)

        # 3) pobierz wszystkie panele wyników
        resp = self.dev.query("MEASure?").strip()
        panels = resp.split('/')
        last = panels[-1]
        parts = [p.strip() for p in last.split(',')]

        if len(parts) < 4:
            raise RuntimeError(f"Niepełne dane z MEASure?: {resp!r}")

        # poprawne przypisanie wartości:
        phase = float(parts[1])
        cp    = float(parts[2])
        d     = float(parts[3])
        rp    = float(parts[4])

        return {
            'Phase': phase,
            'Cp':    cp,
            'D':     d,
            'Rp':    rp,
        }
