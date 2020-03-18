import xlwings as xw
from tkinter import *
from tkinter import messagebox
from PIL import ImageTk
import sched, time

def do_something():
      

    
    Total_Runs = wss.range('B2').options(numbers=int).value
    Overs = wss.range('C2').options(numbers=int).value
    Balls = wss.range('D2').options(numbers=int).value
    Wickets = wss.range('E2').options(numbers=int).value
    Message_Bottom = wss.range('M2').options(numbers=int).value
    Message_B = wss.range('BN13').options(numbers=int).value
    Batsman_1 = wss.range('Z2').options(numbers=int).value
    Runs_1 = wss.range('AA2').options(numbers=int).value
    Balls_1 = wss.range('AB2').options(numbers=int).value
    Now_Playing_1 = wss.range('AC2').options(numbers=int).value    
    Batsman_2 = wss.range('AF2').options(numbers=int).value
    Runs_2 = wss.range('AG2').options(numbers=int).value
    Balls_2 = wss.range('AH2').options(numbers=int).value
    Now_Playing_2 = wss.range('AI2').options(numbers=int).value
    Bowler = wss.range('AT2').options(numbers=int).value
    Wickets_B = wss.range('AX2').options(numbers=int).value
    Runs_B = wss.range('AU2').options(numbers=int).value
    Overs_B = wss.range('AV2').options(numbers=int).value
    Balls_B = wss.range('AW2').options(numbers=int).value
    

                 
    canvas.itemconfig(canvas_id1,text=Total_Runs, font=('Roboto 30 bold'))
    canvas.itemconfig(canvas_id2, text=Overs, font=('Roboto 20 bold'))
    canvas.itemconfig(canvas_id3, text=Balls, font=('Roboto 20 bold'))
    canvas.itemconfig(canvas_id4, text=Wickets, font=('Roboto 30 bold'))
    canvas.itemconfig(canvas_id5, text=Message_Bottom, font=('Roboto 10 bold'))
    canvas.itemconfig(canvas_id6, text=Message_B, font=('Roboto 12 bold'))
    canvas.itemconfig(canvas_id7, text=Batsman_1, font=('Roboto 18 bold'))
    canvas.itemconfig(canvas_id8, text=Runs_1, font=('Roboto 18 bold'))
    canvas.itemconfig(canvas_id9, text=Balls_1, font=('Roboto 12 bold'))
    canvas.itemconfig(canvas_id10, text=Batsman_2, font=('Roboto 18 bold'))
    canvas.itemconfig(canvas_id11, text=Runs_2, font=('Roboto 18 bold'))
    canvas.itemconfig(canvas_id12, text=Balls_2, font=('Roboto 12 bold'))
    canvas.itemconfig(canvas_id13, text=Bowler, font=('Roboto 18 bold'))
    canvas.itemconfig(canvas_id14, text=Wickets_B, font=('Roboto 18 bold'))
    canvas.itemconfig(canvas_id15, text=Runs_B, font=('Roboto 18 bold'))
    canvas.itemconfig(canvas_id16, text=Overs_B, font=('Roboto 15 bold'))
    canvas.itemconfig(canvas_id17, text=Balls_B, font=('Roboto 15 bold'))
    canvas.itemconfig(canvas_id18, text=Now_Playing_1, font=('Roboto 18 bold'))
    canvas.itemconfig(canvas_id19, text=Now_Playing_2, font=('Roboto 18 bold'))
  
    print(Total_Runs)
    print(Overs)
    print('Balls :',Balls)
    print(Wickets)
    print(Message_Bottom)
    print(Batsman_1)
    print(Runs_1)
    print(Balls_1)
    print(Batsman_2)
    print(Runs_2)
    print(Balls_2)
    print(Bowler)
    print(Wickets_B)
    print(Runs_B)
    print(Overs_B)
    print(Balls_B)
    print(Message_B)
    s.enter(0.1, 1, do_something, ())
    master.after(100, do_something)
   



wbt = xw.Book('batting 1st')
wss = wbt.sheets[0]
s = sched.scheduler(time.time, time.sleep)
master = Tk()

#width, height = Image.open(image.png).size
###############
canvas = Canvas(master, width="1920", height="1080", bg="#00ff00")
canvas.pack()

image = ImageTk.PhotoImage(file='1080.png', width="1920", height="1080")
canvas.create_image(960, 540, image=image)


s.enter(0.1, 1, do_something, ())
master.after(100, do_something)
################################### 

#Total Runs
canvas_id1 = canvas.create_text(957, 965, anchor="ne", justify="right")
#Overs
canvas_id2 = canvas.create_text(1099, 1008, anchor="se")
#Balls               
canvas_id3 = canvas.create_text(1113, 1008, anchor="sw")
#Wickets
canvas_id4 = canvas.create_text(1006, 965, anchor="nw")
#Message Bottom
canvas_id5 = canvas.create_text(960, 1021, anchor="n")
#Message B
canvas_id6 = canvas.create_text(1258, 1000, anchor="nw")
#Batsman 1
canvas_id7 = canvas.create_text(289, 964, anchor="nw")#regular daanna bold noda
#Runs 1
canvas_id8 = canvas.create_text(654, 964, anchor="ne", justify="right")
#Balls 1
canvas_id9 = canvas.create_text(661, 971, anchor="nw")#regular daanna bold noda
#Batsman 2
canvas_id10 = canvas.create_text(289, 1000, anchor="nw")
#Runs 2
canvas_id11 = canvas.create_text(654, 1000, anchor="ne", justify="right")
#Balls 2
canvas_id12 = canvas.create_text(661, 1007, anchor="nw")
#Baller
canvas_id13 = canvas.create_text(1258, 964, anchor="nw")
#Wickets Bowler
canvas_id14 = canvas.create_text(1494, 964, anchor="ne", justify="right")
#Runs Baller
canvas_id15 = canvas.create_text(1512, 964, anchor="nw")
#Overs Baller
canvas_id16 = canvas.create_text(1578, 972, anchor="nw")
#Balls Baller
canvas_id17 = canvas.create_text(1584, 980, anchor="ne", justify="right")
#Now Playing 1
canvas_id18 = canvas.create_text(272, 964, anchor="nw")
#Now Playing 2s
canvas_id19 = canvas.create_text(272, 1000, anchor="nw")  

###########
#s.run()
master.mainloop

  #top.mainloop()   
  



