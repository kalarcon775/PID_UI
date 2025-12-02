# test_tc08_interface.py
import time
from tc08_interface import TC08Interface

logger = TC08Interface()

try:
    while True:
        temps = logger.read()
        print(temps)
        time.sleep(1.0)
except KeyboardInterrupt:
    pass
finally:
    logger.close()
