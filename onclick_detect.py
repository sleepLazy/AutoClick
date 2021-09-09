import pyWinhook as pyhook
import pythoncom
from datetime import datetime
import pyautogui
import time

mouse_clicks=[]
hm = pyhook.HookManager()
ctrl=False

class click:
    def __init__(self,x,y,time,drag):
        self.x=x
        self.y=y 
        self.time=time
        self.drag=drag

# release hooks
def release():
    hm.UnhookMouse()
    hm.UnhookKeyboard()

# detect on click events and save 
def onclick(event):
    global mouse_clicks
    global ctrl
    mouse_clicks.append(click(event.Position[0],event.Position[1],datetime.fromtimestamp(event.Time/1e3),ctrl))
    ctrl=False
    return True

# detect keyboard events
def onKeyboardEvent(event):
    global ctrl
    # detect esc to end recording
    if(event.Ascii==27):
        release()
        loop=int(input("Enter times to loop: "))
        print("Automate Starting in 5 second...")        
        auto_click(loop)
    # detect right CTRL to record the next two click as a drag event (start and ending position)
    elif(event.KeyID==162):
        if(ctrl):
            ctrl=False
        else:
            ctrl=True
    return True

#start repeating click events  
def auto_click(loop):
    for j in range(loop):
        time.sleep(2)
        if(not mouse_clicks):
            print("No Mouse Clicks are recorded!")
        else:
            for i in range(0,len(mouse_clicks)):
                if(i != 0):
                    time.sleep((mouse_clicks[i].time-mouse_clicks[i-1].time).total_seconds())
                if(mouse_clicks[i].drag):
                   pyautogui.moveTo(mouse_clicks[i].x,mouse_clicks[i].y)
                   i+=1
                   pyautogui.dragTo(mouse_clicks[i].x,mouse_clicks[i].y,0.5,button='left')
                else:
                    pyautogui.click(mouse_clicks[i].x,mouse_clicks[i].y)
    print("Progress completed!")

if __name__ == '__main__':
    try:
        print ("Detect Starting in 5 sec...") 
        time.sleep(2)
        print("Start Detecting")
        hm.MouseAllButtonsDown=onclick
        hm.HookMouse()
        hm.KeyDown=onKeyboardEvent
        hm.HookKeyboard()
        pythoncom.PumpMessages()
    except:
        release()
        print("Program Ended")