
from guizero import App,Text,PushButton,TextBox
from impe import Impe

def main():

    def funzione():
        c=z1.value
        print(c)
        a=Impe(z1.value,z2.value)
        risultato=a.calcola()

        
        lx.value=str(risultato['lx'])

        cx.value=str(risultato['cx'])

        commento.value=risultato['tipo circuito']


    app=App(title='hello',layout='grid')

    z1_label=Text(app,text='y1',grid=[0,0],align='left')
    z1=TextBox(app,grid=[1,0],width=20,align='left')

    z2_label=Text(app,text='y2',grid=[0,1],align='left')
    z2=TextBox(app,grid=[1,1],width=20,align='left')


    b1=PushButton(app,text='submit',command=funzione,grid=[1,3])

    result_label_1 = Text(app, text="Result_cl:", grid=[0,4], align="left")
    lx = TextBox(app, grid=[1,4], align="left", width=30)

    result_label_2 = Text(app, text="Result_cx:", grid=[0,5], align="left")
    cx = TextBox(app, grid=[1,5], align="left", width=30)

    result_label_3 = Text(app, text="Commento:", grid=[0,6], align="left")
    commento = TextBox(app, grid=[1,6], align="left", width=30)

    



    app.display()


