
# Генерация случайного числа в диапазоне от 0 до 10
import random
number = random.randint(0,10)

# Использование f-строк
print(f"Значение = {number}")
print(f"{number:.2f}")

# Поиск расположения модуля
import inspect
print(inspect.getfile(random))


tmp = [1,2,3,4,5]
tst = list(map(str, tmp))
print(tst)
