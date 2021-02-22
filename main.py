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
        # Create Main Window
        self.main = self.master
        self.main.title("Rabalancing Stocks")
        self.main.geometry("1024x768+100+100")
        
        # Menu Creation
        # File Menu
        self.menu = Menu(self.main)
        self.menu_file = Menu(self.menu, tearoff=0)
        # Default menu disable since there is no user data
        self.menu_file.add_command(label="Update Market Values", command=self.update_data, state="disable") 
        self.menu_file.add_separator()
        self.menu_file.add_command(label="Load User Data", command=self.load_user_data)
        self.menu_file.add_separator()
        self.menu_file.add_command(label="Save to Excel", state="disable") # 비활성화
        self.menu_file.add_separator()
        self.menu_file.add_command(label="Exit", command=self.quit)
        self.menu.add_cascade(label="File", menu=self.menu_file)

        # Strategies Menu
        self.s_var = IntVar()
        self.menu_strategies = Menu(self.menu, tearoff=0)
        self.menu_strategies.add_radiobutton(label="Equal weight", variable=self.s_var, value=1)
        self.menu_strategies.add_radiobutton(label="Market Cap", variable=self.s_var, value=2)
        self.menu_strategies.add_radiobutton(label="P/E Ratio", variable=self.s_var, value=3)
        # Default Strategies menu disable since there is no data to calculate
        self.menu.add_cascade(label="Strategies", menu=self.menu_strategies, state="disable") 

        # Button to print out Strategies data
        self.button_rebalancing = Button(self.main, text='Rebalancing', command=self.rebalancing, state="disable")
        self.button_rebalancing.pack(padx=1, pady=1, fill='x', side='top')

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
        # To test initial balance fixed to $100,000
        invest_amount = 100_000
        number_of_stocks = len(self.new_df.index)
        invest_amount_per_stock = invest_amount/number_of_stocks
        
        for i, price in enumerate(self.new_df['Current Price']):
            temp_series = pd.Series([invest_amount_per_stock/price], index=['Equal Rebalance'])
            temp_series = self.new_df.iloc[i].append(temp_series)
            self.equal_df = self.equal_df.append(temp_series, ignore_index=True)
        self.equal_df = self.equal_df[['Symbol', 'Description', 'Sector', 'Quantity', 'Equal Rebalance', 'Cost Basis Per Share', 'Market Cap', 'Current Price', 'P/E Ratio(PER)', 'P/B Ratio(PBR)']]
        self.pt = Table(self.display, dataframe=self.equal_df, showtoolbar=False, showstatusbar=False)
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
            # Keep only first row if there is duplications
            # self.df.drop_duplicates(subset=['Symbol'], inplace=True)
            # self.df.set_index('Symbol', inplace=True)
            self.df = self.df.head()
            self.pt = Table(self.display, dataframe=self.df, showtoolbar=False, showstatusbar=False)
            self.pt.autoResizeColumns()
            self.pt.show()
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
                    'Description': self.tickers.loc[self.tickers['Symbol'] == symbol]['Security'].item(),
                    'Sector': self.tickers.loc[self.tickers['Symbol'] == symbol]['GICS Sector'].item(),
                    'Cost Basis Per Share': self.df.loc[self.df['Symbol'] == symbol]['Cost Basis Per Share'].item(),
                    'Quantity': self.df.loc[self.df['Symbol'] == symbol]['Quantity'].item(),
                    'Market Cap': data_from_yahoo['marketCap'],
                    'Current Price': data_from_yahoo['regularMarketPrice'],
                    'P/E Ratio(PER)': data_from_yahoo['trailingPE'] if 'trailingPE' in data_from_yahoo.keys() else None,
                    'P/B Ratio(PBR)': data_from_yahoo['priceToBook'] if 'priceToBook' in data_from_yahoo.keys() else None}
                    )
                self.new_df = self.new_df.append(temp, ignore_index=True)
        # Test line to verify loading status
        # print("Loading is completed")
        
        # Reorder Columns
        self.new_df = self.new_df[['Symbol', 'Description', 'Sector', 'Quantity', 'Cost Basis Per Share', 'Market Cap', 'Current Price', 'P/E Ratio(PER)', 'P/B Ratio(PBR)']]
        
        # first_column = self.new_df.pop('Symbol')
        # self.new_df.insert(0, 'Symbol', first_column)
        # self.pt.destroy()
        # self.pt.clearTable()
        self.pt = Table(self.display, dataframe=self.new_df, showtoolbar=False, showstatusbar=False)
        self.pt.autoResizeColumns()
        self.pt.show()
        self.pt.redraw()
        # Update menu status
        self.menu.entryconfig("Strategies", state="normal") 
        self.menu_file.entryconfig("Save to Excel", state="normal")
        self.button_rebalancing.config(state="normal")

app = RebalancingApp()
app.mainloop()