import math
import cmath

class Impe:

    f0=100000000
    omega=None
    risultato={}
    
 
    def __init__(self, y1, y2,f0=100000000,calcolo_eseguito=False):
        self.y1=complex(y1)
        self.y2=complex(y2)
        self.f0=f0
        #self.omega=2*(math.pi)*self.f0
        self.lx=None
        self.cx=None
        self.calcolo_eseguto=calcolo_eseguito
     
    def calcola(self):

        cx=None
        lx=None

        if 1/self.y1.real<1/self.y2.real:
            print('il circuito è serie parallelo')
            self.risultato['tipo circuito']='il circuito è serie parallelo'

            z1=1/(self.y1)

            rs=float(z1.real)*1000
            rp=float(1/self.y2.real)*1000

            x1=z1.imag
            qs=math.sqrt((rp-rs)/rs)

            print('qs',qs)

            self.omega=2*(math.pi)*self.f0

             
            cs=1/(qs*self.omega*rs)

            print('cs',cs)


            cp=cs*(qs*qs/(1+qs*qs))

            print('cp:',cp)
           
            d=self.omega*cp
            
            print('omega*cp:',d)
            print('y2.imag/1000:',self.y2.imag/1000)

            bx=(self.y2.imag)/1000-(self.omega*cp)

            print('bx:',bx)
            

            #print(cp)
            if bx>0:
                self.risultato['cx']=bx/self.omega
                self.risultato['lx']='non previsto'
                print('cx_finale:',self.risultato['cx'])
                self.calcolo_eseguto=True

                
                
            else:
                    
                self.risultato['lx']=1/(abs(bx)*self.omega)
                self.risultato['cx']='non previsto'
                self.calcolo_eseguito=True

            print(print('risultato',self.risultato))

            return self.risultato
                
            
        else:
            
            print('il circuito è parallelo serie')
            rp=1/g1
            rs=z2.real
            self.risultato['lx']='non calcolato'
            self.risultato['cx']='non calcolato1'


        return self.risultato
            

    def stampa(self):
        if self.calcolo_eseguto==True:
            print('cx : {0}  and lx : {1:1.5E}'.format(self.risultato['cx'],self.risultato['lx']))
           