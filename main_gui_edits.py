from customtkinter import *
from CTkMessagebox import *
from PIL import Image
import os
import sys
from aptracker_edits import APTracker
import subprocess
import threading

tracker = APTracker("Test")

tracker.folder_checker()

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

set_appearance_mode("Light")

app_font = ("Arial",13,'normal')
label_font = ("Arial",16,'bold')
header_font = ("Arial",19,'bold')

####Main Window
root  =  CTk()  # create left_panel window
root.title("QSRx Amplitude & Phase Tracking")

root.iconbitmap(resource_path("star-32.ico"))

logos = CTkFrame(root,width=600,  height=200)
logos.pack(side='top',fill='x',expand=True,padx=5,pady=5)

synergy = CTkImage(light_image=Image.open(resource_path('kas-logo.png')),
	dark_image=Image.open(resource_path('kas-logo.png')),
	size=(250,72)) # WidthxHeight

bel = CTkImage(light_image=Image.open(resource_path('BEL-Logo-PNG.png')),
	dark_image=Image.open(resource_path('BEL-Logo-PNG.png')),
	size=(250,97)) # WidthxHeight

syn_logo = CTkLabel(logos, text="", image=synergy)
syn_logo.pack(side='left',padx=50,pady=10)

bel_logo = CTkLabel(logos, text="", image=bel)
bel_logo.pack(side='right',padx=50,pady=10)

width = 160
#width2 = 145
width3 = 85
width4 = 125

height2 = 45
height4 = 27
height = (3*height4+10)/2


width_q = 80
width_big = 3*width_q+20

frame0 = CTkFrame(root,width=600,  height=200)
frame0.pack(side='top',fill='x',expand=True,padx=5,pady=5)

CTkLabel(frame0,text="Instrument Control",font=header_font).pack(side='top',fill='x',expand=True,padx=5,pady=5)

sfb1_sl = StringVar(root,value='123')
sfb2_sl = StringVar(root,value='123')
sfb3_sl = StringVar(root,value='123')
hmrx_sl = StringVar(root,value='123')
qsrx_sl = StringVar(root,value='123')

instr_frame = CTkFrame(frame0)
instr_frame.pack(side='top',fill='x',expand=True,padx=5,pady=5)

instruments_ip_control_frame = CTkFrame(instr_frame)
instruments_ip_control_frame.pack(side="left",fill='x',expand=True,padx=5,pady=5)

instr_label = CTkLabel(instruments_ip_control_frame,text="Instruments",font=label_font)
instr_label.pack(padx=5,pady=5)

instr_status = CTkFrame(instruments_ip_control_frame)
instr_status.pack(side='left',fill='both',expand=True,padx=5,pady=5)

serial_frame = CTkFrame(instr_status)
serial_frame.pack(side='left',fill='both',expand=True,padx=5,pady=5)

CTkLabel(serial_frame,text="SFB1 Sl. No.",font=app_font).grid(row=0,column=1,padx=5,pady=5)
CTkLabel(serial_frame,text="SFB2 Sl. No.",font=app_font).grid(row=0,column=2,padx=5,pady=5)
CTkLabel(serial_frame,text="SFB3 Sl. No.",font=app_font).grid(row=0,column=3,padx=5,pady=5)
CTkLabel(serial_frame,text="QSRx Sl. No.",font=app_font).grid(row=0,column=4,padx=5,pady=5)
CTkLabel(serial_frame,text="HMRx Sl. No.",font=app_font).grid(row=0,column=5,padx=5,pady=5)
CTkEntry(serial_frame,textvariable=sfb1_sl,width=width3).grid(row=1,column=1,padx=7,pady=5)
CTkEntry(serial_frame,textvariable=sfb2_sl,width=width3).grid(row=1,column=2,padx=7,pady=5)
CTkEntry(serial_frame,textvariable=sfb3_sl,width=width3).grid(row=1,column=3,padx=7,pady=5)
CTkEntry(serial_frame,textvariable=qsrx_sl,width=width3).grid(row=1,column=4,padx=7,pady=5)
CTkEntry(serial_frame,textvariable=hmrx_sl,width=width3).grid(row=1,column=5,padx=7,pady=5)

instruments = {
            "Signal Generator": "192.168.10.3",
            "Network Analyzer": "192.168.10.10",
            "Power Supply": "192.168.10.11"
        }

def check_connectivity():

    for instrument,ip_add in instruments.items():
    # Check connectivity using ping
        #is_connected = test_ping(ip_add)
        is_connected = ping_ip_address(ip_add)

        if is_connected:
            status_lights[instrument].configure(text_color="green")
            #messagebox.showinfo("Success", f"{instrument} is connected.")
        else:
            status_lights[instrument].configure(text_color="red")
            #messagebox.showerror("Error", f"{instrument} is not connected.")
        
    return

def abort_button():
    tracker.stop_task()
    update_label_to_idle()
    return

instruments_ip = CTkFrame(instr_status)
instruments_ip.pack(side='left',fill='both',expand=True,padx=5,pady=5)

status_lights = {}

def ping_ip_address(ip_address):
    try:
        # Execute the ping command
        output = subprocess.run(["ping", "-n", "1", ip_address], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        # Check if the ping was successful
        return "Destination host unreachable" not in output.stdout.decode('utf-8') and output.returncode == 0
    except Exception as e:
        print(f"Error pinging {ip_address}: {e}")
        return False

i=0
for instrument, ip_address in instruments.items():
    
    label = CTkLabel(instruments_ip, text=instrument)
    label.grid(row=i,column=0,padx=5)

    ip_label = CTkLabel(instruments_ip, text=ip_address)
    ip_label.grid(row=i,column=1,padx=5)

    status_light = CTkLabel(instruments_ip, text="●", text_color="red")
    status_light.grid(row=i,column=2,padx=5)
    status_lights[instrument] = status_light

    i+=1

CTkButton(instruments_ip, text="Self Test", font=app_font,command=check_connectivity,fg_color="green",height=65,width=80).grid(row=0,column=3,rowspan=3,padx=5,pady=5)

check_connectivity()

ui_status = CTkFrame(instr_frame)
ui_status.pack(side='left',fill='both',expand=True,padx=5,pady=5)

tracking_label = CTkLabel(ui_status,text="Tracking Status",font=label_font)
tracking_label.pack(fill='both',expand=True,padx=5,pady=5)

tracking_frame = CTkFrame(ui_status)
tracking_frame.pack(fill='both',expand=True,padx=5,pady=5)

status_label = CTkLabel(tracking_frame,text="\nIdle",justify="center")
status_label.pack(side="top",fill="both",expand=True,padx=5,pady=5)

CTkButton(tracking_frame,text='Abort',font=app_font,command=abort_button,fg_color='red',height=40).pack(side="top",fill="x",expand=True,padx=5,pady=5)

###SFB###

frame1 = CTkFrame(root,width=600,  height=200)
frame1.pack(side='top',fill='x',expand=True,padx=5,pady=5)

CTkLabel(frame1,text="SFB Tracking",font=header_font).pack(side='top',fill='x',expand=True,padx=5,pady=5)

sfb_content = CTkFrame(frame1,width=600,  height=200)
sfb_content.pack(side='top',fill='x',expand=True,padx=5,pady=5)

ps_frame = CTkFrame(sfb_content)
ps_frame.pack(side='left',fill='x',expand=True,padx=5,pady=5)


sfb_frame = CTkFrame(sfb_content)
sfb_frame.pack(side='left',fill='both',expand=True,padx=5,pady=5)

CTkLabel(sfb_frame,text="QSRx SFB",font=label_font).pack(side='top',fill='x',expand=True,padx=5,pady=5)

hmr_frame = CTkFrame(sfb_content)
hmr_frame.pack(side='left',fill='both',expand=True,padx=5,pady=5)

CTkLabel(hmr_frame,text="HMRx SFB",font=label_font).pack(side='top',fill='x',expand=True,padx=5,pady=5)

frame2 = CTkFrame(root,width=600,  height=200)
frame2.pack(side='top',fill='both',expand=True,padx=5,pady=5)


CTkLabel(frame2,text="QSRx Tracking",font=header_font).pack(side='top',fill='x',expand=True,padx=5,pady=5)
qsr_frame = CTkFrame(frame2)
qsr_frame.pack(side='top',fill='both',expand=True,padx=5,pady=5)


sfb_buttons = CTkFrame(sfb_frame)
sfb_buttons.pack(side='top',fill='x',expand=True,padx=5,pady=5)

CTkLabel(ps_frame,text="Controls",font=label_font).pack(side='top',fill='x',expand=True,padx=5,pady=5)
c_frame1 = CTkFrame(ps_frame)
c_frame1.pack(side='top',fill='both',expand=True,padx=5,pady=5)

CTkButton(c_frame1,text='Turn On PS',font=app_font,command=tracker.submit_ps1_sfb,fg_color='green',width=width4,height=height4).grid(row=0,column=0,padx=5,pady=5)
CTkButton(c_frame1,text='Turn Off PS',font=app_font,command=tracker.turn_off_ps1,fg_color='green',width=width4,height=height4).grid(row=1,column=0,padx=5,pady=5)
CTkButton(c_frame1,text='Data Trace Off',font=app_font,command=tracker.data_trace_off,fg_color='green',width=width4,height=height4).grid(row=2,column=0,padx=5,pady=5)

tracker.sfb1_sl = sfb1_sl.get()
tracker.sfb2_sl = sfb2_sl.get()
tracker.sfb3_sl = sfb3_sl.get()
tracker.hmrx_sl = hmrx_sl.get()
tracker.qsrx_sl = qsrx_sl.get()

def update_label_to_idle():
    status_label.configure(text="\nIdle")
    return

def tracking(key):
    tracking_algo = {1:[tracker.track_sfb1,"SFB1 (RF)"],
                 2:[tracker.track_sfb2,"SFB2 (RF)"],
                 3:[tracker.track_sfb3,"SFB3 (RF)"],
                 4:[tracker.track_sfb1_bite,"SFB1 (BITE)"],
                 5:[tracker.track_sfb2_bite,"SFB2 (BITE)"],
                 6:[tracker.track_sfb3_bite,"SFB3 (BITE)"],
                 7:[tracker.track_qsr,"QSRx (Standalone)"],
                 8:[tracker.track_qsr_band1,"QSRx (Band1)"],
                 9:[tracker.track_qsr_band2,"QSRx (Band2)"],
                 10:[tracker.track_qsr_band3,"QSRx (Band3)"],
                 11:[tracker.track_qsr_int_rf,"QSRx Int. (RF)"],
                 12:[tracker.track_qsr_sfb1_rf,"QSRx Int. SFB1(RF)"],
                 13:[tracker.track_qsr_sfb2_rf,"QSRx Int. SFB2(RF)"],
                 14:[tracker.track_qsr_sfb3_rf,"QSRx Int. SFB3(RF)"],
                 15:[tracker.track_qsr_int_bite,"QSRx Int. (BITE)"],
                 16:[tracker.track_qsr_sfb1_bite,"QSRx Int. SFB1(BITE)"],
                 17:[tracker.track_qsr_sfb2_bite,"QSRx Int. SFB2(BITE)"],
                 18:[tracker.track_qsr_sfb3_bite,"QSRx Int. SFB3(BITE)"],
                 19:[tracker.track_hmr1,"HMR SFB Low (RF)"],
                 20:[tracker.track_hmr1_bite,"HMR SFB Low (BITE)"],
                 21:[tracker.track_hmr2,"HMR SFB High (RF)"],
                 22:[tracker.track_hmr2_bite,"HMR SFB High (BITE)"]}

    tracker.sfb1_sl = sfb1_sl.get()
    tracker.sfb2_sl = sfb2_sl.get()
    tracker.sfb3_sl = sfb3_sl.get()
    tracker.hmrx_sl = hmrx_sl.get()
    tracker.qsrx_sl = qsrx_sl.get()

    status_label.configure(text="Tracking\n %s" % (tracking_algo[key][1]))

    tracking_algo[key][0]()
    
    tracker.stop_event.set()
    #CTkMessagebox(title='Instrument Control', message="Tracking Completed", option_1='OK',icon_size=(40,40),width=500,justify='centre',wraplength=500,font=("Arial",15))
    update_label_to_idle()
    return

def track_thread(key):
    tracker.stop_event.clear()
    task_thread = threading.Thread(target=tracking,args=(key,))
    task_thread.start()

    return


'''SFB Buttons'''

CTkButton(sfb_buttons,text='Track SFB1',font=app_font,command=lambda: track_thread(1),width=width,height=height).grid(row=3,column=0,padx=5,pady=5)
CTkButton(sfb_buttons,text='Track SFB2',font=app_font,command=lambda: track_thread(2),width=width,height=height).grid(row=3,column=1,padx=5,pady=5)
CTkButton(sfb_buttons,text='Track SFB3',font=app_font,command=lambda: track_thread(3),width=width,height=height).grid(row=3,column=2,padx=5,pady=5)
#CTkButton(sfb_frame,text='Track HMRx',font=app_font,command=track_hmrx,width=width,height=height).grid(row=1,column=2,padx=5,pady=5)

CTkButton(sfb_buttons,text='Track SFB1 \n(BITE)',font=app_font,command=lambda: track_thread(4),width=width,height=height).grid(row=4,column=0,padx=5,pady=5)
CTkButton(sfb_buttons,text='Track SFB2 \n(BITE)',font=app_font,command=lambda: track_thread(5),width=width,height=height).grid(row=4,column=1,padx=5,pady=5)
CTkButton(sfb_buttons,text='Track SFB3 \n(BITE)',font=app_font,command=lambda: track_thread(6),width=width,height=height).grid(row=4,column=2,padx=5,pady=5)

'''QSRX Controls'''

q_frame1 = CTkFrame(qsr_frame)
q_frame1.pack(side='left',fill='both',expand=True,padx=5,pady=5)

CTkLabel(q_frame1,text="Controls",font=label_font).pack(side='top',fill='x',expand=True,padx=5,pady=5)
q_controls = CTkFrame(q_frame1)
q_controls.pack(side='left',fill='both',expand=True,padx=5,pady=5)

CTkButton(q_controls,text='Turn On PS',font=app_font,command=tracker.submit_ps1_qsr,fg_color='green',width=width4,height=height4).grid(row=0,column=0,padx=5,pady=5)
CTkButton(q_controls,text='Turn Off PS',font=app_font,command=tracker.turn_off_ps1,fg_color='green',width=width4,height=height4).grid(row=1,column=0,padx=5,pady=5)
CTkButton(q_controls,text='Data Trace Off',font=app_font,command=tracker.data_trace_off_qsr,fg_color='green',width=width4,height=height4).grid(row=2,column=0,padx=5,pady=5)

'''QSRX Buttons'''

q_frame2 = CTkFrame(qsr_frame)
q_frame2.pack(side='left',fill='both',expand=True,padx=5,pady=5)

CTkLabel(q_frame2,text="QSRx Standalone",font=label_font).pack(side='top',fill='both',expand=True,padx=5,pady=5)
qsrx_standalone = CTkFrame(q_frame2)
qsrx_standalone.pack(side='left',fill='both',expand=True,padx=5,pady=5)

CTkButton(qsrx_standalone,text='QSRx Standalone',font=app_font,command=lambda: track_thread(7),width=width_big,height=height).grid(row=0,column=0,columnspan=3,padx=5,pady=5)
CTkButton(qsrx_standalone,text='Band1',font=app_font,command=lambda: track_thread(8),width=width_q,height=height).grid(row=1,column=0,padx=5,pady=5)
CTkButton(qsrx_standalone,text='Band2',font=app_font,command=lambda: track_thread(9),width=width_q,height=height).grid(row=1,column=1,padx=5,pady=5)
CTkButton(qsrx_standalone,text='Band3',font=app_font,command=lambda: track_thread(10),width=width_q,height=height).grid(row=1,column=2,padx=5,pady=5)

q_frame3 = CTkFrame(qsr_frame)
q_frame3.pack(side='left',fill='both',expand=True,padx=5,pady=5)

CTkLabel(q_frame3,text="QSRx Integrated (RF)",font=label_font).pack(side='top',fill='both',expand=True,padx=5,pady=5)
qsrx_integrated = CTkFrame(q_frame3)
qsrx_integrated.pack(side='left',fill='both',expand=True,padx=5,pady=5)

CTkButton(qsrx_integrated,text='QSRx Int. (RF)',font=app_font,command=lambda: track_thread(11),width=width_big,height=height).grid(row=0,column=0,columnspan=3,padx=5,pady=5)
CTkButton(qsrx_integrated,text='SFB1',font=app_font,command=lambda: track_thread(12),width=width_q,height=height).grid(row=1,column=0,padx=5,pady=5)
CTkButton(qsrx_integrated,text='SFB2',font=app_font,command=lambda: track_thread(13),width=width_q,height=height).grid(row=1,column=1,padx=5,pady=5)
CTkButton(qsrx_integrated,text='SFB3',font=app_font,command=lambda: track_thread(14),width=width_q,height=height).grid(row=1,column=2,padx=5,pady=5)

q_frame4 = CTkFrame(qsr_frame)
q_frame4.pack(side='left',fill='both',expand=True,padx=5,pady=5)

CTkLabel(q_frame4,text="QSRx Integrated (BITE)",font=label_font).pack(side='top',fill='both',expand=True,padx=5,pady=5)
qsrx_integrated_b = CTkFrame(q_frame4)
qsrx_integrated_b.pack(side='left',fill='both',expand=True,padx=5,pady=5)

CTkButton(qsrx_integrated_b,text='QSRx Int. (BITE)',font=app_font,command=lambda: track_thread(15),width=width_big,height=height).grid(row=0,column=0,columnspan=3,padx=5,pady=5)
CTkButton(qsrx_integrated_b,text='SFB1',font=app_font,command=lambda: track_thread(16),width=width_q,height=height).grid(row=1,column=0,padx=5,pady=5)
CTkButton(qsrx_integrated_b,text='SFB2',font=app_font,command=lambda: track_thread(17),width=width_q,height=height).grid(row=1,column=1,padx=5,pady=5)
CTkButton(qsrx_integrated_b,text='SFB3',font=app_font,command=lambda: track_thread(18),width=width_q,height=height).grid(row=1,column=2,padx=5,pady=5)

'''HMRx Buttons'''

hmr_buttons = CTkFrame(hmr_frame)
hmr_buttons.pack(side='top',fill='x',expand=True,padx=5,pady=5)

CTkButton(hmr_buttons,text='Track HMRx low',font=app_font,command=lambda: track_thread(19),width=width,height=height).grid(row=0,column=0,padx=5,pady=5)
CTkButton(hmr_buttons,text='Track HMRx low \n(BITE)',font=app_font,command=lambda: track_thread(20),width=width,height=height).grid(row=1,column=0,padx=5,pady=5)
CTkButton(hmr_buttons,text='Track HMRx high',font=app_font,command=lambda: track_thread(21),width=width,height=height).grid(row=0,column=1,padx=5,pady=5)
CTkButton(hmr_buttons,text='Track HMRx high \n(BITE)',font=app_font,command=lambda: track_thread(22),width=width,height=height).grid(row=1,column=1,padx=5,pady=5)

root.resizable(False,False)
root.mainloop()