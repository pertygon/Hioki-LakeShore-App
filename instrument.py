import pyvisa
import time
import random

class MockLakeshore335:
    """Symulator kontrolera temperatury."""
    def __init__(self, *args, **kwargs):
        self._target = 300.0

    def set_temperature(self, T):
        print(f"[MOCK Lake] ustawiam temperaturę na {T:.2f} K")
        self._target = T

    def get_temperature(self):
        # symuluj powolne zbliżanie do T
        # zwróć wartość losowo odchyloną o ±0.1 K
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
        # zwróć słownik z przykładami realnych‐wyglądających wartości
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
        rm = pyvisa.ResourceManager('')
        self.dev = rm.open_resource(resource_name)
        self.dev.write_termination = '\n'
        self.dev.read_termination = '\n'

    def set_temperature(self, T):
        # komenda zależna od interfejsu (GPIB/RS232) i SCPI Lakeshore
        self.dev.write(f"SETP 1,{T:.2f}")

    def get_temperature(self):
        return float(self.dev.query("KRDG? 1"))

class Hioki3536:
    def __init__(self, resource_name):
        rm = pyvisa.ResourceManager('')
        self.dev = rm.open_resource(resource_name)
        self.dev.write_termination = '\n'
        self.dev.read_termination = '\n'

    def set_frequency(self, freq_hz):
        self.dev.write(f"FREQ {freq_hz:.0f}")

    def measure_all(self):
        data = self.dev.query("READ?").split(',')
        # kolejność odpowiada dokumentacji: Phase, Cp, D, Rp
        return {
            'Phase': float(data[0]),
            'Cp':    float(data[1]),
            'D':     float(data[2]),
            'Rp':    float(data[3]),
        }
