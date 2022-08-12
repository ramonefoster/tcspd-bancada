import serial
import threading
import serial.tools.list_ports
import time
from PyQt5 import QtTest

class AxisControll():
    def __init__(self, device, port, baund):
        self.porta = port
        self.baundRate = baund
        self.port=self.porta 
        #AH or DEC 
        self.device = device  
        self.time_out_device = 2    

        self.result = self.com_ports()

        self.error_device = False

        if self.porta in self.result:
            self.ser = serial.Serial(
            port=self.porta,
            baudrate=self.baundRate,
            timeout=self.time_out_device,
            writeTimeout=1
            )
            self.ser.close()
            if self.ser.is_open == False:
                try: 
                    self.ser.open()
                    self.ser.reset_output_buffer()
                    self.ser.reset_input_buffer()
                    self.error_device = False
                except Exception as e:
                    self.error_device = True                    
        else:
            print('Cannot connect to: ', self.porta)
            self.error_device = True

    def close_port(self):
        try:
            if self.ser.is_open:
                self.ser.cancel_write()
                self.ser.cancel_read()
                self.ser.reset_output_buffer()
                self.ser.reset_input_buffer()                
                print("Closing Port")
                self.ser.close()
        except Exception as e:
            print("Close Port error: ", e)

    def com_ports(self):
            self.list = serial.tools.list_ports.comports()
            self.connected = []
            for element in self.list:
                self.connected.append(element.device)

            return(self.connected)

    def progStatus(self):    
        if self.error_device:
            return("+0 00 00.00 *0000000000000000")
        else:        
            try:  
                ack = self.write_cmd(self.device+" PROG STATUS\r")

                if len(ack) > 2:
                    return(ack)    
                else:
                    print("ProgStatus bug")
                    return("+0 00 00.00 *0000000000000000")
            except Exception as e:
                print(e)
                return("+0 00 00.00 *0000000000000000")

    def mover_rap(self, position):
        if not self.error_device:           
            ret = 'ACK' in self.write_cmd(self.device+" EIXO MOVER_RAP = " + str(position) + "\r")
            if ret:
                stat = True
            else:
                stat = False
            return stat 

    def mover_rel(self, position):
        if not self.error_device:
            ret = 'ACK' in self.write_cmd(self.device+" EIXO MOVER_REL = " + str(position) + "\r")
            if ret:
                stat = True
            else:
                stat = False
            return stat   

    def girar_vel(self, velocidade):
        if not self.error_device:
            ret = 'ACK' in self.write_cmd(self.device+" EIXO GIRAR_VEL = " + str(velocidade) + "\r")
            if ret:
                stat = True
            else:
                stat = False
            return stat                  

    def prog_error(self):
        if not self.error_device:
            ret = self.write_cmd(self.device+" PROG ERROS\r")
            if ret:
                stat = True
            else:
                stat = False
            return ret 
                    
    def prog_parar(self):
        if not self.error_device:
            ret = 'ACK' in self.write_cmd(self.device+" PROG PARAR\r")
            if ret:
                stat = True
            else:
                stat = False
            return stat

    def sideral_ligar(self):
        if not self.error_device:
            ret = 'ACK' in self.write_cmd(self.device+" SIDERAL LIGAR\r")
            if ret:
                stat = True
            else:
                stat = False
            return stat

    def sideral_desligar(self):
        if not self.error_device:
            ret = 'ACK' in self.write_cmd(self.device+" SIDERAL DESLIGAR\r")
            if ret:
                stat = True
            else:
                stat = False
            return stat  

    def reset(self):
        if not self.error_device:
            ret = 'ACK' in self.write_cmd(self.device+" PROG RESET\r")
            if ret:
                stat = True
            else:
                stat = False
            return stat   
            
    def write_cmd(self, cmd):
        if not self.error_device and self.ser.is_open:
            try:
                self.ser.flush()
                QtTest.QTest.qWait(250)
                self.ser.reset_output_buffer()
                self.ser.reset_input_buffer()
                print(cmd)
                self.ser.write(cmd.encode())
                ack = ''
                t0 = time.time()
                while '\r' not in ack:
                    ack += self.ser.read().decode()
                    if (time.time() - t0) > self.time_out_device:
                    # if (len(ack) < 3) or 'NAK' in ack:
                        self.ser.reset_input_buffer()
                        self.ser.reset_output_buffer()
                        print("serial timeout", ack)
                        
                    #     break
                        return ack
                print("statbuf: "+ack)    
                return ack.replace('\r', '')
            except Exception as e:
                print("error writing COM: ", e)
                self.ser.cancel_read()
                self.ser.cancel_write()
                return('NAK')
        else:
            return('NAK')


axis_thread = threading.Thread(target = AxisControll, args=[0, 0, 0])
axis_thread.start()