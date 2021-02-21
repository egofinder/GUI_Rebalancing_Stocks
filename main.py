import pandas as pd
from tkinter import *
import tkinter.ttk as ttk # Library for Prograssive Bar library
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
        self.menu_file.add_command(label="Exit", command=self.main.quit)
        self.menu.add_cascade(label="File", menu=self.menu_file)

        # Strategies Menu
        self.menu_strategies = Menu(self.menu, tearoff=0)
        self.menu_strategies.add_radiobutton(label="Equal weight")
        self.menu_strategies.add_radiobutton(label="Market Cap")
        self.menu_strategies.add_radiobutton(label="P/E Ratio")
        # Default Strategies menu disable since there is no data to calculate
        self.menu.add_cascade(label="Strategies", menu=self.menu_strategies, state="disable") 

        # Add menu bar to main Window
        self.main.config(menu=self.menu)
        
        # Display Frame to top at Window
        self.display = Frame(self.main)
        self.display.pack(padx=1, pady=1, fill=BOTH, side='top', expand=1)
        
        # Display Progressbar to bottom at Window
        self.p_var2 = DoubleVar()
        self.progressbar2 = ttk.Progressbar(self.main, maximum=100, variable=self.p_var2)
        self.progressbar2.pack(padx=1, pady=1, side='bottom', fill='x')
        

# Functions
    def load_market_data(self) -> 'pandas.DataFrame':
        self.tickers = tickers_sp500(True)
        self.tickers.set_index('Symbol', inplace=True)
        
    def load_user_data(self) -> 'pandas.DataFrame':
        _file = filedialog.askopenfilename(title="Choose csv file...",  \
            filetypes=(("csv", "*.csv"),("All", "*.*")), initialdir=r"C:\Users\baekh\OneDrive\Desktop\Python")
        try:
            self.df = pd.read_csv(_file, index_col=False)
            self.df = self.df.drop(columns=['Account Name/Number'])
            # Keep only first row if there is duplications
            self.df.drop_duplicates(subset=['Symbol'], inplace=True)
            self.df.set_index('Symbol', inplace=True)
            # self.df = self.df.head(5)
            self.pt = Table(self.display, dataframe=self.df, showtoolbar=False, showstatusbar=False)
            self.pt.autoResizeColumns()
            self.pt.show()
            self.pt.redraw()
            # Update Menu status
            self.menu_file.entryconfig("Update Market Values", state="normal")
        except:
            pass
             
    def update_data(self) -> 'pandas.DataFrame':
        self.load_market_data()
        for i, symbol in enumerate(self.df.index):
            # Variable for display status of Progressive Bar
            self.p_var2.set(100/len(self.df.index)*(i+1))
            self.progressbar2.update()

            if symbol in self.tickers.index:
                # Loading live market data from yahoo finance
                data_from_yahoo = get_quote_data(symbol)
                # print(data_from_yahoo.keys())
                
                temp = pd.Series(
                    {'Symbol': symbol,
                    'Company': self.tickers.loc[symbol]['Security'],
                    'Cost Basis Per Share': self.df.loc[symbol]['Cost Basis Per Share'],
                    'Quantity': self.df.loc[symbol]['Quantity'],
                    'Market Cap': data_from_yahoo['marketCap'], 
                    'Current Price': data_from_yahoo['regularMarketPrice'],
                    'P/E Ratio(PER)': data_from_yahoo['trailingPE'] if 'trailingPE' in data_from_yahoo.keys() else None,
                    'P/B Ratio(PBR)': data_from_yahoo['priceToBook'] if 'priceToBook' in data_from_yahoo.keys() else None}
                    )
                self.new_df = self.new_df.append(temp, ignore_index=True)
        # Test line to verify loading status
        # print("Loading is completed")
        first_column = self.new_df.pop('Symbol')
        self.new_df.insert(0, 'Symbol', first_column)
        # self.pt.destroy()
        # self.pt.clearTable()
        self.pt = Table(self.display, dataframe=self.new_df, showtoolbar=False, showstatusbar=False)
        self.pt.autoResizeColumns()
        self.pt.show()
        self.pt.redraw()
        # Update menu status
        self.menu.entryconfig("Strategies", state="normal") 
        self.menu_file.entryconfig("Save to Excel", state="normal")

app = RebalancingApp()
app.mainloop()