
from selenium import webdriver
import numpy as np , pandas as pd
import matplotlib.pyplot as plt
from selenium.webdriver.chrome.options import Options
import time
import glob
import os
from matplotlib import gridspec
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from tkinter import *
from tkinter import Frame, messagebox
from tkinter.ttk import Style

class Crawler:
    def __init__(self):
        self.chr_path =  'C:\\Users\\Khashi\\Desktop\\job\\google_webdrive\\92\\chromedriver' 
        self.dl_path = "C:\\Users\\Khashi\\Desktop\\py3\\giti\\stocks\\repo"                      
        self.url = 'https://www.fipiran.com/DataService/ISIndex'
        self.stock_name = None
        self.year = None
        self.bot_option = None
        self.subject = []
        self.last_date = 0
        self.last_period = 0
        self.options()
    
    def options(self ):
        global driver
        options = webdriver.ChromeOptions()
        prefs = {"download.default_directory" : self.dl_path }
        options.add_experimental_option("prefs",prefs)
        # options.add_argument("--headless")
        self.bot_option = options
        
    def err_run(self , err=0):
        try:
            self.run()
            self.bar_plot()
        except ValueError:
            err = 'nothing was found'
        print('---> error ' , err)
        return err

    def run(self, stock, year):
        self.stock_name = stock
        self.year = year

        driver = webdriver.Chrome(chrome_options = self.bot_option , executable_path = self.chr_path) 
        driver.get(self.url)
        time.sleep(2)
        WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,"//button[contains(text(),'متوجه شدم')]"))).click()
        driver.find_element_by_id('symbolparaIS').send_keys(stock)
        time.sleep(0.5)
        driver.find_element_by_id('year').send_keys(year)
        time.sleep(0.5)
        driver.find_element_by_id('gobutton').click()
        time.sleep(2)
        driver.quit()
 
    def save_excel(self):
        list_of_files = glob.glob( self.dl_path+"\\*") 
        latest_file = max(list_of_files, key=os.path.getctime)
        data = pd.read_html(latest_file)
        data[0].to_excel(self.dl_path+"\\{}.xlsx".format(self.stock_name))
        os.remove(latest_file)

        print('--> save excel done!')

    def bar_plot(self , plot = True):        
        if os.path.exists(self.dl_path+"\\{}.xlsx".format(self.stock_name)):
            print('xlsx exist ')
            data = pd.read_excel(self.dl_path+"\\{}.xlsx".format(self.stock_name))
        else:
            list_of_files = glob.glob( self.dl_path+"\\*.xls") 
            latest_file = max(list_of_files, key=os.path.getctime)
            data = pd.read_html(latest_file)
            data = data[0]

        fin_yr = data.loc[0 , 'FinanceYear']
        fig = plt.figure(figsize = (5, 5))
        sty = self.stock_name[::-1]
        fig.suptitle('{} financial net income/ financial_year {} '.format(sty , fin_yr), fontsize=14, fontweight='bold')
        ax0 = plt.subplot(1,1,1)
        ax0.set_xlabel('period')
        ax0.set_ylabel('NetIncome Million rial')

        for k , i in enumerate(data.priod):
            # print(k , i)
            if i == 3:
                ax0.bar(data.loc[k , ['priod']].values , data.loc[k , ['NetIncome']].values , label = '3_month')
                period , pub_date , net_in =str(data.loc[k , ['priod']].values[0]) , str(data.loc[k , ['publishDate']].values[0]) , int(data.loc[k , ['NetIncome']].values[0])
                txt = period +' month financial statement' + 'published in ' + pub_date + ' and net income is ' + f"{net_in:,}" + ' Million Rial'
                self.subject.append(txt)
            if i == 6:
                ax0.bar(data.loc[k , ['priod']].values , data.loc[k , ['NetIncome']].values , label ='6_month')
                period , pub_date , net_in =str(data.loc[k , ['priod']].values[0]) , str(data.loc[k , ['publishDate']].values[0]) , int(data.loc[k , ['NetIncome']].values[0])
                txt = period +' month financial statement' + 'published in ' + pub_date + ' and net income is ' + f"{net_in:,}" + ' Million Rial'
                self.subject.append(txt)
            if i == 9:
                ax0.bar(data.loc[k , ['priod']].values , data.loc[k , ['NetIncome']].values , label = '9_month')  
                period , pub_date , net_in =str(data.loc[k , ['priod']].values[0]) , str(data.loc[k , ['publishDate']].values[0]) , int(data.loc[k , ['NetIncome']].values[0])
                txt = period +' month financial statement' + 'published in ' + pub_date + ' and net income is ' + f"{net_in:,}" + ' Million Rial'
                self.subject.append(txt)
            if i ==12:
                ax0.bar(data.loc[k , ['priod']].values , data.loc[k , ['NetIncome']].values , label = '12_month')   
                period , pub_date , net_in =str(data.loc[k , ['priod']].values[0]) , str(data.loc[k , ['publishDate']].values[0]) , int(data.loc[k , ['NetIncome']].values[0])
                txt = period +' month financial statement' + 'published in ' + pub_date + ' and net income is ' + f"{net_in:,}" + ' Million Rial'
                self.subject.append(txt)

        self.last_date = data.loc[: , ['publishDate']].values.max()
        x = data.loc[data['publishDate'] == self.last_date]
        self.last_period = x['priod'].values
        ####
        max_dur = max(self.last_period)
        net_in = x['NetIncome'].loc[x['priod'] == max_dur].values[0]
        net = format(net_in, ',d')
        ####
        ax0.axhline(color = "red" , label = "zero_line")   
        ax0.grid()
        ax0.legend()
        if plot == True:
            plt.tight_layout()
            plt.show()
        if plot == False:
            return  self.last_date , max_dur , net
          
    def text_plot(self):
        fig = plt.figure(figsize = ( 12 , 15)) 
        ax0 = plt.subplot2grid((15,4), (1,0), colspan=4 , rowspan = 12)
        ax1 = plt.subplot2grid((15,4) , (13,0) , colspan=4 , rowspan=2)
        stock = self.stock_name[::-1]
        fig.suptitle('__{}__'.format(stock), fontsize=14, fontweight='bold' , color = 'red')
        ax0.set_ylabel(' dates of Financial statement ')
        lenght = len(self.subject)
        # print('lenght of statement is ' , lenght)
        for k , i in enumerate(self.subject):
            print(i, '  ' , type(i))
            ax0.text(x=0.2 , y= 0.02 + float(k/lenght) , s =  i  ,fontsize = 10)
        
        ax1.set_ylabel(' last Financial statement ')
        text = ' last Financial statement release in ' + str(self.last_date) + ' which period is {}'.format(str(self.last_period)) + ' month'
        # print(text)
        ax1.text(x=0.2 , y=0.5 , s = text , fontsize = 15 )
        plt.tight_layout()
        ax0.set_xticks([])
        ax0.set_yticklabels([])
        ax1.set_xticks([])
        ax1.set_yticklabels([])
        plt.tight_layout()
        plt.show()

class GUI(Crawler, Frame):
    def __init__(self, back):
        Frame.__init__(self)
        Crawler.__init__(self)
        # self.year = None
        self.back = back
        self.configure(bg = 'gray37')
        self.create_widgets()
    
    def plot(self):
        self.run(self.entry.get() , self.year_entry.get())
        messagebox.showinfo(message="---- get data successfully  ----")
        self.bar_plot()

    def save_xlsx(self):
        self.save_excel()
        messagebox.showinfo(message="---- convert .xls to .xlsx <DONE>----")

    def exit(self):
        self.back.destroy()

    def create_widgets(self):
        self.style = Style()
        self.pack(fill = BOTH, expand=True)
        self.style.theme_use('alt') 
        self.columnconfigure(0 , weight = 5)    
        self.columnconfigure(1 , weight = 5)

        self.entry = Entry( self , width=20)
        self.entry.grid(column=0 , row=0 , padx = 10 , pady = 20)
        # self.gr_but = Button(self, width = )
        self.year_entry = Entry( self , width=20)
        self.year_entry.grid(column=0 , row=1, padx = 10 , pady = 20)

        self.nput_label = Label(self, width=10 , text = 'stock_name' , bg = 'khaki3').grid(column=1 , row = 0 , sticky='NSEW' , pady = 20 , padx = 10)
        self.nput_label = Label(self, width=10 , text = 'financial_year' , bg = 'khaki3').grid(column=1 , row = 1 , sticky='NSEW' , pady = 20 , padx = 10)
        self.save_but = Button(self , command = self.save_excel , text = 'save_result' , bd = '5' , bg = 'khaki3' , width = 10).place(x = 115 , y = 200 )
        self.plot_but = Button(self , command= self.plot , bd = '5' , text = 'plot' , bg = 'khaki3' ).grid(row = 2 , column = 0 , sticky='NSEW' , padx = 10 , pady = 10)
        self.txt_plot = Button(self , command = self.text_plot , text = 'text_plot' , bd = '5 ' , bg = 'khaki3').grid(row = 2 , column=1 , sticky='NSEW' , pady = 10 , padx = 10)
        self.quit = Button(self , text = 'Quit' , command= self.exit , bd = '5' , bg = 'khaki3' , width = 10 ).place(x  = 115 , y = 240)

def main():
    canvas = Tk()
    canvas.geometry("340x300+300+300")
    canvas.title('--- fipiran crawler ---')
    GUI(back=canvas)
    canvas.mainloop()

if __name__ == '__main__':
    main()
