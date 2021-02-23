import pandas as pd
from tkinter import *
from tkinter import messagebox # Library for messagebox
import tkinter.ttk as ttk # Library for Prograssive Bar
from tkinter import filedialog # Library for CSV File opening
from pandastable import Table # Library to disply Pandas DataFrame to Tkinter
from yahoo_fin.stock_info import tickers_sp500
from yahoo_fin.stock_info import get_quote_data

class RebalancingApp(Frame):
    def __init__(self, parent=None):
        self.parent = parent
        self.df = pd.DataFrame()
        self.new_df = pd.DataFrame()
        self.tickers = pd.DataFrame()
        
        super().__init__()
        self.invested_value = StringVar(value="0")
        self.add_invest_value = StringVar(value="0")
        # Create Main Window
        self.main = self.master
        self.main.title("Rabalancing Stocks")
        self.main.geometry("1024x768+100+100")
        self.top_frame = Frame(self.main)
        self.top_frame.pack(side ='top')
        
        # Menu Creation
        # File Menu
        self.menu = Menu(self.main)
        self.menu_file = Menu(self.menu, tearoff=0, font="* 12")
        # Default menu disable since there is no user data
        self.menu_file.add_command(label="Update Market Values", command=self.update_data, state="disable") 
        self.menu_file.add_separator()
        self.menu_file.add_command(label="Load User Data", command=self.load_user_data)
        self.menu_file.add_separator()
        self.menu_file.add_command(label="Save to Excel", state="disable")
        self.menu_file.add_separator()
        self.menu_file.add_command(label="Exit", command=self.quit)
        self.menu.add_cascade(label="File", menu=self.menu_file)

        # Strategies Menu
        self.s_var = IntVar(value=0)
        self.menu_strategies = Menu(self.menu, tearoff=0, font="* 12")
        self.menu_strategies.add_radiobutton(label="Equal weight", variable=self.s_var, value=1)
        self.menu_strategies.add_radiobutton(label="Market Cap", variable=self.s_var, value=2, state="disable")
        self.menu_strategies.add_radiobutton(label="P/E Ratio", variable=self.s_var, value=3, state="disable")
        # Default Strategies menu disable since there is no data to calculate
        self.menu.add_cascade(label="Strategies", menu=self.menu_strategies, state="disable") 

        # Button to print out Strategies data
        self.invested_label_name = Label(self.top_frame, bg='#76EE00', font="* 20", text="Invesed Money($): ")
        self.invested_label_name.grid(row=0, column=0, padx=1, pady=1)
        self.invested_label = Label(self.top_frame, bg='#FFF8DC', font="* 20", width = 13, textvariable=self.invested_value)
        self.invested_label.grid(row=0, column=1, padx=1, pady=1)
        
        self.add_invest_name = Label(self.top_frame, bg='#76EE00', font="* 20", text="Extra Money($): ")
        self.add_invest_name.grid(row=0, column=2, padx=1, pady=1)
        self.add_invest_amount = Entry(self.top_frame, bg='#FFF8DC', font="* 20", width = 13, textvariable=self.add_invest_value)
        self.add_invest_amount.grid(row=0, column=3, padx=1, pady=1)
        
        self.button_rebalancing = Button(self.top_frame, text='Rebalancing', font="* 15", command=self.rebalancing, state="disable")
        self.button_rebalancing.grid(row=0, column=4, padx=0, pady=0)     

        # Add menu bar to main Window
        self.main.config(menu=self.menu)
        
        # Display Frame to top at Window
        self.display = Frame(self.main)
        self.display.pack(padx=1, pady=1, fill=BOTH, side='top', expand=1)
        
        # Display Progressbar to bottom at Window
        self.p_var = DoubleVar()
        self.progressbar = ttk.Progressbar(self.main, maximum=100, variable=self.p_var)
        self.progressbar.pack(padx=1, pady=1, side='bottom', fill='x')
        
# Functions
    # Window X Button overide
    # def on_closing():
    #     if messagebox.askokcancel("Quit", "Do you want to quit?"):
    #         root.destroy()

    def quit(self):
        if messagebox.askyesno("Exit","Do you want to quit the application?"):
            self.main.quit()

    def rebalancing(self):
        if self.s_var.get() == 0:
            messagebox.showerror("Error","Please, select your strategy!")
        
        elif self.s_var.get() == 1:
            self.equal_rebalancing()
            
        elif self.s_var.get() == 2:
            self.market_cap_rebalancing()
            
        elif self.s_var.get() == 3:
            self.pe_ratio_rebalancing()
        
        else:
            messagebox.showerror("Critical","Something wrong!!!!!!!!")
        
# Rabalancing Functions    
    def equal_rebalancing(self):
        self.equal_df = pd.DataFrame()
        # I have to fix when user input text instead intgeter.
        if self.add_invest_value.get() == "0":
            invest_amount = self.invested_temp_sum
        else:
            invest_amount = int(self.invested_temp_sum) + int(self.add_invest_amount.get())

        number_of_stocks = len(self.new_df.index)
        invest_amount_per_stock = invest_amount/number_of_stocks
        
        for i, price in enumerate(self.new_df['Current Price']):
            temp_series = pd.Series([invest_amount_per_stock/price], index=['Equal Rebalance'])
            temp_series = self.new_df.iloc[i].append(temp_series)
            self.equal_df = self.equal_df.append(temp_series, ignore_index=True)
        self.equal_df = self.equal_df[['Symbol', 'Description', 'Sector', 'Quantity', 'Equal Rebalance', 'Cost Basis Per Share', 'Market Cap', 'Current Price', 'P/E Ratio(PER)', 'P/B Ratio(PBR)']]
        self.pt = Table(self.display, dataframe=self.equal_df, showtoolbar=False, showstatusbar=False, editable=False)
        self.pt.autoResizeColumns()
        self.pt.show()
        self.pt.redraw()
        
    def market_cap_rebalancing(self):
        print('Market Cap')
        
    def pe_ratio_rebalancing(self):
        print('P/E Ratio')  
               
    def load_market_data(self):
        self.tickers = tickers_sp500(True)
        # self.tickers.set_index('Symbol', inplace=True)
        
    def load_user_data(self):
        _file = filedialog.askopenfilename(title="Choose csv file...",  \
            filetypes=(("csv", "*.csv"),("All", "*.*")), initialdir=r"C:\Users\baekh\OneDrive\Desktop\Python")
        try:
            self.df = pd.read_csv(_file, index_col=False)
            self.df = self.df.drop(columns=['Account Name/Number'])
            self.df = self.df.rename(columns={"Current Value":"Current Price"})
            # Keep only first row if there is duplications
            # self.df.drop_duplicates(subset=['Symbol'], inplace=True)
            self.df = self.df.head()           
            self.pt = Table(self.display, dataframe=self.df, showtoolbar=False, showstatusbar=False, editable=False)
            self.pt.show()
            self.pt.autoResizeColumns()
            self.pt.redraw()
            # Update Menu status
            self.menu_file.entryconfig("Update Market Values", state="normal")
        except:
            pass
             
    def update_data(self):
        self.load_market_data()
        for i, symbol in enumerate(self.df['Symbol']):
            # Variable for display status of Progressive Bar
            self.p_var.set(100/len(self.df.index)*(i+1))
            self.progressbar.update()

            if symbol in self.tickers['Symbol'].values:
                # Loading live market data from yahoo finance
                data_from_yahoo = get_quote_data(symbol)
                # print(data_from_yahoo.keys())
                temp = pd.Series(
                    {'Symbol': symbol,
                    'Description': self.tickers.loc[self.tickers['Symbol'] == symbol]['Security'].item().strip(),
                    'Sector': self.tickers.loc[self.tickers['Symbol'] == symbol]['GICS Sector'].item().strip(),
                    'Cost Basis Per Share': self.df.loc[self.df['Symbol'] == symbol]['Cost Basis Per Share'].item().strip(),
                    'Quantity': self.df.loc[self.df['Symbol'] == symbol]['Quantity'].item(),
                    'Market Cap': data_from_yahoo['marketCap'],
                    'Current Price': data_from_yahoo['regularMarketPrice'],
                    'P/E Ratio(PER)': data_from_yahoo['trailingPE'] if 'trailingPE' in data_from_yahoo.keys() else None,
                    'P/B Ratio(PBR)': data_from_yahoo['priceToBook'] if 'priceToBook' in data_from_yahoo.keys() else None}
                    )
                self.new_df = self.new_df.append(temp, ignore_index=True)
        # print("Loading is completed")
        # Reorder Columns
        self.new_df = self.new_df[['Symbol', 'Description', 'Sector', 'Quantity', 'Cost Basis Per Share', 'Market Cap', 'Current Price', 'P/E Ratio(PER)', 'P/B Ratio(PBR)']]
        # first_column = self.new_df.pop('Symbol')
        # self.new_df.insert(0, 'Symbol', first_column)
        # self.pt.destroy()
        # self.pt.clearTable()
        self.invested_temp_sum = 0
        for i in range(len(self.new_df)):
            self.invested_temp_sum += self.new_df['Quantity'][i] * self.new_df['Current Price'][i]
        self.invested_temp_sum = round(self.invested_temp_sum, 3) 
        self.invested_value.set(str(self.invested_temp_sum))
        self.pt = Table(self.display, dataframe=self.new_df, showtoolbar=False, showstatusbar=False, editable=False)
        self.pt.show()
        self.pt.autoResizeColumns()
        self.pt.redraw()
        
        # Update menu status
        self.menu.entryconfig("Strategies", state="normal") 
        self.menu_file.entryconfig("Save to Excel", state="normal")
        self.button_rebalancing.config(state="normal")

app = RebalancingApp()
app.mainloop()