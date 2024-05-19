import mouse
import pyautogui
import keyboard
import tkinter as tk 
import time
import random
import math
import win32api
import win32gui
import win32con
import pywintypes
import pygetwindow
import pyautogui

max_move_speed = 1000
stopstartkey = 'PgUp'
quitkey = 'End'
catchkey = 'shift'

screen_width, screen_height = pyautogui.size()
centerx = screen_width / 2
centery = screen_height / 2
bounce = 0.95
x, y =  mouse.get_position()
CanToggle = True
Running = True
Runall = True
gettime = time.perf_counter()
xmomentum = 0
ymomentum = 0
movex = 0
movey = 0
friction = 0.98
slow = 2
pull = 4
gravity = 0.7
castrope = False
ropehealth = 255
ropedecay = 0.05
ropeheal = 1
broken = False
mendingthreashold = 100
stamina = 50
windowstrength = 100
windowlostgrip = False
winddowdropchancepersec = 0.75
fps = 60
holdtime = 0
shaketime = 0
shakepower = 0
shakemultiplyer=2
shakewindowintensity = 0
checkedwindows = False

def is_user_grabbing_window():
    mouse_x, mouse_y = win32gui.GetCursorPos()
    active_window = pygetwindow.getActiveWindow()
    if active_window is not None and not active_window.isMinimized:  
        # Check if the active window is not the desktop window or the shell window
        if active_window.title != "Program Manager" and active_window.title != "Shell_TrayWnd":
            window_rect = active_window.topleft
            window_width, window_height = active_window.size
            if (window_rect[0] <= mouse_x <= window_rect[0] + window_width) and \
               (window_rect[1] <= mouse_y <= window_rect[1] + window_height):
                if win32api.GetKeyState(0x01) < 0:
                    return True
    return False

def move_active_window(dx, dy):
    try:
        active_window = pygetwindow.getActiveWindow()
        if active_window is not None:
            new_x = active_window.left + dx
            new_y = active_window.top + dy
            active_window.moveTo(new_x, new_y)
    except Exception:
        pass 

def shake_windows(intensity, window_effects):
    for window_effect in window_effects:
        window_effect.shake_window(intensity)
        window_effect.shake()

class WindowShakeEffect:
    def __init__(self, window):
        self.window = window
        self.loopspeed = 20
        self.shaking = False
        self.initial_position = (self.window.left, self.window.top)
        self.gettime = time.perf_counter()
        self.screen_width, self.screen_height = win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1)
        self.intensity = 0

    def update_position(self):
        self.initial_position = (self.window.left, self.window.top)

    def shake_window(self, intensity):
        self.intensity = intensity
        if not self.shaking:
            self.shaking = True
            self.gettime = time.perf_counter()

    def shake(self):
        if self.shaking:
            if time.perf_counter() - self.gettime <= 0.3:
                progress = (time.perf_counter() - self.gettime) / 0.3  # Duration of 0.3 seconds
                new_x = self.initial_position[0] + random.randint(-self.intensity, self.intensity) * (1 - progress)
                new_y = self.initial_position[1] + random.randint(-self.intensity, self.intensity) * (1 - progress)
                new_x += random.randint(-self.intensity, self.intensity)
                new_y += random.randint(-self.intensity, self.intensity)
                
                # Ensure the new position is within the screen bounds with shake window intensity
                try:
                    new_x = max(0 - self.intensity, min(new_x, self.screen_width - self.window.width + self.intensity))
                    new_y = max(0 - self.intensity, min(new_y, self.screen_height - self.window.height + self.intensity))
                    self.window.moveTo(round(new_x), round(new_y))
                except pygetwindow.PyGetWindowException:
                    pass  # Ignore windows we can't move
            else:
                self.shaking = False



        


def draw_rope(x1, y1, x2, y2):
    distance = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
    red = min(255 - int(ropehealth), 255)
    green = min(int(ropehealth), 255)
    color = win32api.RGB(red, green, 0)
    hdc = win32gui.GetDC(0)
    pen = win32gui.CreatePen(win32con.PS_SOLID, 1, color)
    win32gui.SelectObject(hdc, pen)
    win32gui.MoveToEx(hdc, int(x1), int(y1))
    win32gui.LineTo(hdc, int(x2), int(y2))
    win32gui.DeleteObject(pen)
    win32gui.ReleaseDC(0, hdc)

while Runall:
    deltaTime = (time.perf_counter() - gettime)
    gettime = time.perf_counter()
    
    if keyboard.is_pressed(quitkey):
        Runall = False
        
    if Running == True:
        if keyboard.is_pressed(stopstartkey):
            if CanToggle == True:
                Running = False
                CanToggle = False
        else:
            CanToggle = True
        
        x_prev, y_prev = x, y 
        x, y = mouse.get_position()
        
        totalxmovement = x - x_prev
        totalymovement = y - y_prev
        xmomentum += (x - x_prev - round(xmomentum))/slow + (round(movex)-round(movex)/slow)
        ymomentum += (y - y_prev - round(ymomentum))/slow + (round(movey)-round(movey)/slow) + (gravity-gravity/slow) + gravity
        
        if is_user_grabbing_window():
            holdtime += 1
            if holdtime>=windowstrength:
                if random.random() <= winddowdropchancepersec*deltaTime:
                    windowlostgrip = True
            if windowlostgrip:
                move_active_window(totalxmovement,totalymovement)
            elif holdtime<stamina:
                xmomentum *= 0.0
                ymomentum *= 0.0
            elif holdtime<stamina*2: 
                xmomentum *= 0.5
                ymomentum *= 0.5
            else:
                xmomentum *= 0.75
                ymomentum *= 0.75
        elif castrope:
            xmomentum *= 0.99
            ymomentum *= 0.93
            holdtime = 0
            windowlostgrip = False
        else:        
            xmomentum *= friction
            ymomentum *= friction 
            ymomentum += gravity-gravity*friction
            holdtime = 0
            windowlostgrip = False
        
        if keyboard.is_pressed(catchkey) and not broken:
            if not castrope:
                centerx, centery = mouse.get_position()
                castrope = True
            dx = centerx - x
            dy = centery - y
            distance = math.sqrt(dx ** 2 + dy ** 2)
            if distance > 0:
                speed = pull * deltaTime
                movex = dx / distance * min(speed * distance, max_move_speed)
                movey = dy / distance * min(speed * distance, max_move_speed)
            else:
                movey = 0
                movex = 0
            if ropehealth > 0:
                ropehealth -= distance*ropedecay
            else:
                broken = True
            draw_rope(x,y,centerx,centery)
        else:
            if castrope:
                castrope = False
                movey = 0
                movex = 0
            if ropehealth < 255:
                ropehealth += ropeheal
            else:
                ropehealth = 255
            if ropehealth >= mendingthreashold:
                broken = False
        mouse.move(x + round(xmomentum) + round(movex)*castrope, y + round(ymomentum) + round(movex)*castrope )
        tempx, tempy = mouse.get_position()
        if tempx >= screen_width-1 or tempx <= 0:
            if abs(xmomentum) > 3:
                if shakepower < abs(xmomentum)*shakemultiplyer:
                    shakepower = abs(xmomentum)*shakemultiplyer
                    if shakepower>1:
                        shaketime = 1
                    else:
                        shaketime = 0
                xmomentum *= -bounce
                mouse.move(x , y)
        if tempy >= screen_height-1 or tempy <= 0:
            if abs(ymomentum) > 3:
                if shakepower < abs(ymomentum)*shakemultiplyer:
                    shakepower = abs(ymomentum)*shakemultiplyer
                    if shakepower>1:
                        shaketime = 1
                    else:
                        shaketime = 0
            ymomentum *= -bounce
            mouse.move(x , y)
        if shaketime > 0:
            if not checkedwindows:
                windows = pygetwindow.getAllWindows()
                window_effects = [WindowShakeEffect(window) for window in windows]
                checkedwindows = True
            if round(shakepower*shaketime*shaketime) < 1:
                shaketime = 0
            else:
                shakewindowintensity = round(shakepower*shaketime*shaketime)
                shake_windows(round(shakepower*shaketime*shaketime),window_effects)
            shaketime -= deltaTime
        else: 
            shakepower = 0
            if checkedwindows:
                checkedwindows = False
    else:
        screen_width, screen_height = pyautogui.size()
        centerx = screen_width / 2
        centery = screen_height / 2
        x, y =  mouse.get_position()
        if keyboard.is_pressed(stopstartkey):
            if CanToggle == True:
                Running = True
                CanToggle = False
        else:
            CanToggle = True
    time.sleep(1/fps)



