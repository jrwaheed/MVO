
from flask import Flask, redirect, url_for
from flask import render_template, session
from flask import request, redirect

import itertools
from itertools import combinations
from itertools import product
import xlrd
import select as something
import numpy as np
import pandas as pd
from functools import reduce
from bokeh.io import output_notebook, show
from bokeh.plotting import figure, show, output_file, ColumnDataSource
import jinja2
from bokeh.embed import components
from IPython.display import HTML
from bokeh.models import HoverTool, LabelSet,LinearColorMapper
from bokeh.transform import factor_cmap
import gc

Asset = ['UK large cap',
'UK mid cap',
'UK small cap',
'US equities',
'Europe ex UK equities',
'Asia Pacific ex Japan equities',
'Japan equities',
'Emerging market equities',
'Global REITs',
'Global ex UK bonds',
'UK bonds',
'Cash']

SD_Directory = ('C:/Users/for_m_000/AppData/Local/Programs/Python/Python36-32/Asset ClassExpected Return.xlsx')
RET_Directory = ('C:/Users/for_m_000/AppData/Local/Programs/Python/Python36-32/Asset ClassExpected Return.xlsx')
COR_Directory = ('C:/Users/for_m_000/AppData/Local/Programs/Python/Python36-32/Asset Correlation.xlsx')

# Weight Interval Selection ------------------------------------------------------------------------------------------------------------------------------------------
Interval5=[]
Interval10=[]
Interval20=[]

Interval_Dict={ "Interval5":(0,.05,.10,.15,.20,.25,.30,.35,.40,.45,.50,.55,.60,.65,.70,.75,.80,.85,.90,.95), 
                "Interval10":(0,.10,.20,.30,.40,.50,.60,.70,.80,.90), 
                "Interval20":(0,.20,.40,.60,.80)}

Interval_Dict_h={"Interval5": '5% Interval', 
                "Interval10": '10% Interval',
                "Interval20": '20% Interval'}


Real_Weights = []

#HTML------------------------------------------------------------------------------------------------------------------------------------------------

df_RET_Directory = pd.read_csv('C:/Users/for_m_000/AppData/Local/Programs/Python/Python36-32/Asset Expected Return2.csv')
df_COR_Directory = pd.read_csv('C:/Users/for_m_000/AppData/Local/Programs/Python/Python36-32/Asset Correlation.csv')

Asset_Index_Selected =[]
Asset_Index_Html =[]
Risk_Free_Rate = []
Risk_Free_Rate2 =[]
Interval_Post = ""
User_Input_Interval = []

p3 = figure(plot_width = 820, plot_height =540, title="Efficient Frontier")
p3.circle(15, 7, size = 10, line_color = "red", fill_color = "red", fill_alpha =1)
p3.border_fill_color = "white"
p3.xaxis.axis_label = 'Portfolio Standard Deviation (%)'
p3.yaxis.axis_label = 'Portfolio Expectd Return (%)'
script, div = components(p3)




app = Flask(__name__)
@app.route("/", methods = ["GET","POST"])
def inflow():
    global Corr_Indexed
    
    if request.method == 'POST':
        Interval_Post =(request.form.get('Interval_Html'))
        Risk_Free_Rate = request.form.get('RFR_Html')
        Asset_Index_Html = (request.form.getlist('mycheckbox'))

#EARLY SYSTEM BREAKS!!!-------------------------------------------------------------------------------------------------------------------------------
        if request.form.get('RFR_Html') =="":
            return render_template("Template4.html")
        if not Risk_Free_Rate:
            return render_template("Template4.html")
        if len(Asset_Index_Html) < 2:
            return render_template("Template6.html")
        #-------------------------------------------------------------------------------------------------------------------------
        
        return redirect(url_for("user", Interval_Post=Interval_Post, Risk_Free_Rate=Risk_Free_Rate, Asset_Index_Html=Asset_Index_Html))
    
    else:   
        return render_template("Template1.html", 
        Returns_h=[df_RET_Directory.to_html(classes='table')], 
        Correlations_h=[df_COR_Directory.to_html(classes='table')],
        script = script, div = div)  

@app.route("/ <Risk_Free_Rate> / <Interval_Post> / <Asset_Index_Html>")
def user(Risk_Free_Rate, Interval_Post, Asset_Index_Html):
    global Corr_Indexed

    Standard_Deviations = []
    Exp_Rtns = []
    Correlations = []
    
    Asset_Number =1
    Weight =1

    Asset_Index_Selected =[]
    Asset_Number = []
    
    Asset_Index_Html = Asset_Index_Html.replace("'", "")
    Asset_Index_Html = Asset_Index_Html.replace("[", "")
    Asset_Index_Html = Asset_Index_Html.replace("]", "")
    
    Asset_Index_Html = Asset_Index_Html.split(",")

    for i in (Asset_Index_Html):
        Asset_Index_Selected.append(int(i))


    print(type(Asset_Index_Selected))
    print(Asset_Index_Selected)
    

    Asset_Number = len(Asset_Index_Selected)
    print(Asset_Number)
    
    Asset_Selected = list(np.array(Asset)[Asset_Index_Selected])
    print("Assets named:", Asset_Selected)
    
    #LATE SYSTEM BREAK!!!!!-------------------------------------------------------------------------------------------------------------------------
    if Asset_Number > 11 and Interval_Post == "Interval20":
        return render_template("Template3.html")
    if Asset_Number > 8 and Interval_Post == "Interval10":
        return render_template("Template3.html")
    if Asset_Number > 6 and Interval_Post == "Interval5":
        return render_template("Template3.html")
    if request.form.get('RFR_Html') =="":
        return render_template("Template4.html")
    if not Risk_Free_Rate:
        return render_template("Template4.html")
    
    #Risk Free Rate----------------------------------------------------------------------------------------------------------------------------------   
    Risk_Free_Rate2 = (float(Risk_Free_Rate))
    print("Risk Free rate", Risk_Free_Rate2)

    Interval_h = Interval_Dict_h[Interval_Post]

    #def perm(n, Intervals:------------------------------------------------------------------------------------------------------------------------------------------------
    Real_Weights=[]
    for k in list(itertools.product((Interval_Dict[Interval_Post]),repeat=Asset_Number)): 
        if sum(k) == 1:
            (Real_Weights.append(k))
        else:
            continue
    #print(Real_Weights)   


    #Finding the Applicable Standard deviations and return rates, based on the selected assets-----------------------------------------------------------------------
    
    for i in Asset_Index_Selected:
        
        wb =xlrd.open_workbook(SD_Directory)
        sheet = wb.sheet_by_index(0)
        Standard_Deviations.append(sheet.cell_value(i + 1 ,2))
        Exp_Rtns.append(sheet.cell_value(i + 1 ,1))
    #print("Expected Returns", Exp_Rtns)  
    
    #Finding the applicable Correlations based on the selected assets-----------------------------------------------------------------
    Corr_Indexed=[] 
    Corr_Indexed = (list(itertools.combinations(Asset_Index_Selected, 2)))
    #print(Corr_Indexed)

    for i,j in Corr_Indexed:
        wb =xlrd.open_workbook(COR_Directory)
        sheet = wb.sheet_by_index(0)
        Correlations.append(sheet.cell_value(i + 1 ,j + 1))

    #print('Applicable correlations:', Correlations)

    # 'Zipping' or tying each Standard deviation to a row/column in Excel matrix
    SD_Zip = zip(Asset_Index_Selected,Standard_Deviations)
    SD_Indexed_Zipped = list(SD_Zip)
    #print('Zipped Standard Deviations', SD_Indexed_Zipped)

    # 'Zipping' or tying each correlation to a coordinate in Excel matrix
    Corr_Zip = zip(Corr_Indexed,Correlations)
    Correlations_Indexed_UnZipped = list(Corr_Zip)

    #Expanded Zipped Correlations
    Correlations_Indexed_Zipped = [(x,y,z) for (x,y),z in Correlations_Indexed_UnZipped]
    #print('Zipped Correlations, correctly indexed:', Correlations_Indexed_Zipped)

    #Order correlations correctly with Standard deviations
    Portfolio_Variance_End =[]
    Cycle_SD1 =[]
    Cycle_SD2 =[]
    SD1 =[]
    SD2 =[]

    for a,b,c in Correlations_Indexed_Zipped: 
        for d,e in SD_Indexed_Zipped:
            if d == a:
                Cycle_SD1 = e
                SD1.append(Cycle_SD1)
                continue

    for a,b,c in Correlations_Indexed_Zipped: 
        for d,e in SD_Indexed_Zipped:
            if d == b:
                Cycle_SD2 = e
                SD2.append(Cycle_SD2)
                continue        

    SDProd = []
    SDProd = [a * b for a,b in zip(SD1,SD2)]

    #print(SD1)
    #print(SD2)
    #print(Correlations)

    Num_of_Weights = [i for i in range (Asset_Number)]   
    #print('Number of Weights:', Num_of_Weights)

    #Create Pandas Dataframe-----------------------------------------------------------------------------------
    #MVO_Assets = ["Assets", "Standard Deviations 1", "Standard Deviation 2", "Correlations", "Expected Returns"]
    #pd.DataFrame(columns=MVO_Assets)
    df_ASSET = pd.DataFrame({'Assets': Asset_Selected })
    df_RET = pd.DataFrame({'Expected Returns': Exp_Rtns})
    df_RET = df_RET.transpose()
    df_CORR =pd.DataFrame({'Correlations': Correlations})
    df_SD_BEG = pd.DataFrame({'Standard Deviations':Standard_Deviations})
    df_SD_BEG =df_SD_BEG.transpose()
    df_SD_END = pd.DataFrame({'Standard Deviation 1': SD1, 'Standard Deviation 2': SD2,'Product SD': SDProd}) 
    df_WGT = pd.DataFrame(Real_Weights, columns =[Num_of_Weights])



    # Portfolio Return--------------------------------------------------------------------------------------------------------------------

    PORT_RET_SUM =[]
    df_PORT_RET = (df_WGT.mul(df_RET.iloc[0].values))
    PORT_RET_SUM = df_PORT_RET.sum(axis = 1)
    df_PORT_RET['Port Return Summation'] = PORT_RET_SUM

    #---------------------------------------------------------------------------------------------------------------------------------------------


    # Portfolio Variance 1st Half---------------------------------------------------------------------------------------------------------

    PORT_VAR1_SUM =[]
    df_PORT_VAR1 =(df_WGT**2).mul((df_SD_BEG**2).iloc[0].values)
    PORT_VAR1_SUM = df_PORT_VAR1.sum(axis =1)
    df_PORT_VAR1['Port Variance1 Summation'] =PORT_VAR1_SUM
    #print('Portfolio Variance 1:', df_PORT_VAR1)

    # Portfolio Variance 2nd Half----------------------------------------------------------------------------------------------------------

    df_WGT_VAR2_TU = df_WGT.apply(lambda r: list(combinations(r,2)), axis = 1)
    #print(df_WGT_VAR2_TU)

    PORT_VARlist = []

    for tup in df_WGT_VAR2_TU:
        for elem in tup:
            shot = elem[0] * elem[1]
            PORT_VARlist.append(shot)


    #print("Number of rows:", len(PORT_VARlist)) 
    #print("Number of columns:", len(df_CORR))

    PORT_VAR2ROWS = int(len(PORT_VARlist) / len(df_CORR)) 
    PORT_VAR2COL = len(df_CORR)

    PORT_VAR2_WGT = np.array(PORT_VARlist).reshape(PORT_VAR2ROWS , PORT_VAR2COL)    
    df_PORT_VAR2_WGT = pd.DataFrame(PORT_VAR2_WGT)
    #print(PORT_VAR2_WGT)

    # Portfolio Variance 2nd Half, Weights * Correlation * Standard Deviations---------------------------------------------------------------------

    CORR_SD_PROD =[]
    CORR_SD_PROD = [c * d for c,d in zip(SDProd , Correlations)]
    #print(CORR_SD_PROD)

    tester =[]
    tester1 =[]
    for j in PORT_VAR2_WGT:
        q = 2 * j * CORR_SD_PROD
        tester.append(q)
    tester1 = np.array(tester).reshape(PORT_VAR2ROWS , PORT_VAR2COL) 
    #print(tester1)

    w =[]
    PORT_VAR2_SUM = []

    for ele in tester1:
        w = sum(ele)
        PORT_VAR2_SUM.append(w)

    df_PORT_VAR2 = pd.DataFrame(PORT_VAR2_SUM) 
    #print(df_PORT_VAR2)


    # TOTAL Variance and corresponding return ----------------------------------------------------------------------------------------------------------------------------

    PORT_VAR_TOTAL = []
    PORT_VAR_TOTAL = [e + f for e,f in zip(PORT_VAR1_SUM, PORT_VAR2_SUM)]


    PORT_SD_TOTAL =[]
    PORT_SD_TOTAL = np.sqrt(PORT_VAR_TOTAL)
    #print(PORT_SD_TOTAL)

    SHARPE1 = []
    SHARPE1 = np.array(PORT_RET_SUM)
    SHARPE1 = SHARPE1 - Risk_Free_Rate2

    SHARPE2 = []
    SHARPE2 = [g / h for g,h in zip(SHARPE1 , PORT_SD_TOTAL)]


    MATRIX1 = pd.DataFrame({'Portfolio Return': PORT_RET_SUM,'Portfolio Variance':PORT_VAR_TOTAL,'Portfolio Standard Deviation':PORT_SD_TOTAL, 'Sharpe':SHARPE2})


    MATRIX2 = pd.DataFrame(Real_Weights, columns=(Asset_Selected))


    df_FINAL_MATRIX = pd.concat([MATRIX1,MATRIX2], axis=1)
    print("")
    print("Final Matrix")
    print(df_FINAL_MATRIX)

    # Find Max Sharpe Ratio---------------------------------------------------------------------------------------------------------------------------------------------

    OPTIMAL_PORT =[]
    OPTIMAL_PORT = df_FINAL_MATRIX['Sharpe'].idxmax()
    print("")
    print("Optimal Portfolio")

    print(df_FINAL_MATRIX.loc[[OPTIMAL_PORT]])



    # Plot Scatterplot ------------------------------------------------------------------------------------------------------------------------------------------------

    OPTIMAL_PORT_HTML =[]
    OPTIMAL_PORT_HTML = (df_FINAL_MATRIX.loc[[OPTIMAL_PORT]])
    OPTIMAL_PORT_HTML_T = OPTIMAL_PORT_HTML.transpose()
    OPTIMAL_PORT_HTML_T =OPTIMAL_PORT_HTML_T.round(decimals=2)
    print(OPTIMAL_PORT_HTML_T)


    df_MATRIX_DROP = df_FINAL_MATRIX.drop(df_FINAL_MATRIX.index[OPTIMAL_PORT])
    print("")
    print("Matrix with drop")
    print(df_MATRIX_DROP)

    OP_RETURN_HTML = (OPTIMAL_PORT_HTML['Portfolio Return'].iloc[0])
    OP_SD_HTML = (OPTIMAL_PORT_HTML['Portfolio Standard Deviation'].iloc[0])

    Tooltip_Expansion_List = []
    for i in Asset_Selected:
        Tooltip_Expansion_List.append((i,"@{" + i + "}{0.00}"))

    Tooltip_Expansion_Tuple = tuple(Tooltip_Expansion_List)


    source = ColumnDataSource(df_MATRIX_DROP)
    source2 = ColumnDataSource(df_FINAL_MATRIX.loc[[OPTIMAL_PORT]])

    tooltips_pre = ()
    tooltips_pre = ("Portfolio Return","$y"),("Portfolio Standard Deviation","$x"),("Sharpe Ratio",'@Sharpe')

    tooltips = tooltips_pre + Tooltip_Expansion_Tuple 
    tooltips = list(tooltips)

    p = figure(plot_width = 820, plot_height =540,tooltips=tooltips, title="Efficient Frontier")
    p.circle("Portfolio Standard Deviation", "Portfolio Return", source = source, 
    size = 5, line_color = "blue", fill_color = "blue", fill_alpha =0.7)
    p.circle("Portfolio Standard Deviation", "Portfolio Return", source = source2, 
    size = 10, line_color = "red", fill_color = "red", fill_alpha =1)
    p.xaxis.axis_label = 'Portfolio Standard Deviation (%)'
    p.yaxis.axis_label = 'Portfolio Expectd Return (%)'
    script, div = components(p)


    #Render to HTML Template ---------------------------------------------------------------------------------------------------------------
    return render_template("Template2.html",
    portfolio_h=[OPTIMAL_PORT_HTML_T.to_html(classes='table')],
    #titles=OPTIMAL_PORT_HTML.columns.values,
    Returns_h=[df_RET_Directory.to_html(classes='table')], 
    Correlations_h=[df_COR_Directory.to_html(classes='table')],
    script = script, div = div, RFR_h=Risk_Free_Rate, Interval_h=Interval_h,
    Assets_Selected_h = Asset_Selected)

    print("At the back door")
    del Corr_Indexed, df_WGT, Real_Weights, Num_of_Weights
    gc.collect()

@app.route("/garbageman", methods = ["GET"])
def garbageman():
    gc.collect()
    #app.run()  
    return redirect("/")

@app.route("/profile", methods = ["POST","GET"])
def profile():
    if request.method == 'GET':
        return render_template("Template5.html")


if __name__ == "__main__":
    app.run(Debug = True)        