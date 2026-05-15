import pyautogui as pi
from time import sleep

sleep(5)

print(pi.position())

for pos in pi.locateAllOnScreen(r"img\estoqueProdutos.png", confidence=0.8):
    loc = pi.center(pos)
    pi.click(loc)
    print(pi.position())


