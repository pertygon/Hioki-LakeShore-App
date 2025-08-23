# instrument.py

import pyvisa
import time
import random
from pyvisa import constants
from lakeshore import Model335

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
    def close(self):
        pass

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
    """
    Sterownik Lake Shore Model 335 przez lake-shore-python-driver.
    Jeśli instrument nie obsługuje kanału !=1, używa kanału 1.
    """

    def __init__(self, resource_name, baud_rate=57600, timeout=2.0):
        if resource_name.upper().startswith("ASRL"):
            # wyciągam numer portu COM z nazwy VISA
            num = resource_name[4:resource_name.find("::")]
            com_port = f"COM{num}"
            self.dev = Model335(baud_rate, com_port=com_port, timeout=timeout)
        else:
            self.dev = Model335(baud_rate, timeout=timeout)
    def set_temperature(self, T, channel=2):
        """
        Ustawia setpoint T [K] na wyjściu `channel`.
        Jeśli SETP dla podanego kanału zwróci błąd, powtórzy dla kanału 1.
        """
        try:
            self.dev.set_control_setpoint(channel, T)
            self.dev.set_heater_range(channel, 'LOW')
        except Exception:
            # fallback do kanału 1
            self.dev.set_control_setpoint(1, T)
            self.dev.set_heater_range(channel, 'LOW')

    def disable_heater(self,channel=2):
        """
        Ustawia setpoint T [K] na wyjściu `channel`.
        Jeśli SETP dla podanego kanału zwróci błąd, powtórzy dla kanału 1.
        """
        try:
            self.dev.set_heater_range(channel, self.dev.HeaterRange.OFF)
        except Exception:
            self.dev.set_heater_range(channel, self.dev.HeaterRange.OFF)
    def enable_heater(self,channel=2):
        """
        Ustawia setpoint T [K] na wyjściu `channel`.
        Jeśli SETP dla podanego kanału zwróci błąd, powtórzy dla kanału 1.
        """
        try:
            self.dev.set_heater_range(channel, self.dev.HeaterRange.HIGH)
        except Exception:
            self.dev.set_heater_range(channel, self.dev.HeaterRange.HIGH)
    def get_temperature(self, channel=2):
        """
        Odczytuje temperaturę [K] z kanału `channel`.
        Jeśli odczyt z kanału 2 się nie uda, pobiera z kanału 1.
        """
        try:
            temp = self.dev.get_all_kelvin_reading()
            return temp[1]
        except Exception:
            temp = self.dev.get_all_kelvin_reading()
            return temp[0]
    def get_heater_output(self, channel=2):
        """
        Zwraca procent mocy grzałki (0–100%) na zadanym kanale.
        """
        return self.dev.get_heater_output(channel)
    def close(self):
        """Zamknij port COM używany wewnętrznie przez lakeshore.Model335."""
        try:
            self.dev.device_serial.close()
        except Exception:
            pass

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

        for cmd in ("FUNC 'CPD'", "TRIG:SOUR IMM", "FORM:DATA ASCII"):
            try:
                self.dev.write(cmd)
            except:
                pass

    def set_frequency(self, freq_hz):
        self.dev.write(f"FREQ {freq_hz:.0f}")

    def measure_all(self):
        time.sleep(0.05)
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
            pass

        # 3) pobierz wszystkie panele wyników
        resp = self.dev.query("MEASure?").strip()
        print(resp)
        panels = resp.split('/')
        last = panels[-1]
        parts = [p.strip() for p in last.split(',')]
        print(parts)
        if(len(parts) > 4):
            parts = parts[1:]
        if len(parts) < 4:
            raise RuntimeError(f"Niepełne dane z MEASure?: {resp!r}")

        # poprawne przypisanie wartości:
        phase = float(parts[0])
        cp    = float(parts[1])
        d     = float(parts[2])
        rp    = float(parts[3])

        return {
            'Phase': phase,
            'Cp':    cp,
            'D':     d,
            'Rp':    rp,
        }
