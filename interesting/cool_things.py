
# Генерация случайного числа в диапазоне от 0 до 10
import random
number = random.randint(0,10)

# Использование f-строк
print(f"Значение = {number}")
print(f"{number:.2f}")

# Замерить скорость работы куска кода
class Profiler(object):
    def __enter__(self):
        self._startTime = time.time()

    def __exit__(self, type, value, traceback):
        print("Elapsed time: {:.3f} sec".format(time.time() - self._startTime))

with Profiler() as p:
    print("Тут должен быть измеряемый кусок кода")

# Вычисление Наибольшего Общего Делителя
def gcd(a, b):
    while b:
        a, b = b, a % b
    return a



# Поиск расположения модуля
import inspect
print(inspect.getfile(random))