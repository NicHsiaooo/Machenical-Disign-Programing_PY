import pandas as pd
import numpy as np
import math

Force_ = 9800    #(N)
Syt = 325    #(mpa)
SF = 2
#(矩形管尺寸參考 cns7741規格，參考網址: "https://jindesp.com.tw/manual")


#暫時假設
Up_b = 100      #(mm)       #???代處理

#抓取excel資料
def read_excel_file(file_path):
    df = pd.read_excel(file_path)
    print(f"成功讀取 Excel 文件: {file_path}\n")
    return df

def find_index_in_array(array, target_number):
    for index, item in enumerate(array):
        if isinstance(item, dict) and target_number in item.values():
            return index


if __name__ == '__main__':
#讀取拉桿長度尺寸表
    file_path1 = "Length_DataFarme.xlsx"      #假設長度的Excel的路徑
    Read_df = read_excel_file(file_path1)        #該Excel 資料放在變數Read_df裡，資料型態為DataFrame
    Data_dict = Read_df.to_dict(orient='records')
    #print(Data_dict)

#讀取下拉桿截面尺寸表
    file_path2 = "Down_Area_size.xlsx"      
    Read_Down_df = read_excel_file(file_path2)        
    Down_Data_dict = Read_Down_df.to_dict(orient='records')
    #print(Down_Data_dict)
    
#讀取上拉桿截面尺寸表
    file_path3 = "Up_Area_size.xlsx"      
    Read_Up_df = read_excel_file(file_path3)        
    Up_Data_dict = Read_Up_df.to_dict(orient='records')
    #print(Up_Data_dict)

    count = 0
    for Data in Data_dict:
        print(f"項目{count+1}:\n")        #print

        Ld = Data["Down_L"]
        Lu = Data["Up_L"]
        Lb = Data["Base_L"]
        Lw = Data["Wall_L"]
        Aw = Data["Wall_A"]
        Au = Data["Up_A"]

#角度、長度計算(驗算完成)----------------------------------------------------------------------------------------------------------------------------------------

        Assume_C = (Lb**2 + Lw**2)**0.5

        if Assume_C>=Ld+Lu or Ld>=Assume_C+Lu or Lu>=Assume_C+Ld:        #C, Lu, Ld 圍成一個三角形

            count=count+1
            continue

        ag1 = math.acos((Lu**2-Ld**2-Assume_C**2)/(-2*Ld*Assume_C))
        agb = ag1 + math.acos(Lb/Assume_C)
        ag2 = math.acos((Ld**2-Lu**2-Assume_C**2)/(-2*Lu*Assume_C))
        age = ag2+math.acos(Lw/Assume_C)
        Data_dict[count]['A'] = (Au**2+Aw**2-(2*Au*Aw*math.cos(age)))**0.5

        #油壓鋼伸長量要小於其原長的一半
        C0 = math.sqrt(Lb**2+(Lw+2000)**2)
        if Lu>=Ld+C0 or Ld>=Lu+C0 or C0>=Lu+Ld:
            print("0公尺時, 油壓鋼三角形不合")
            count = count+1
            continue

        agm = math.atan(Lb/(Lw+2000))
        agn = math.acos((Lu**2+C0**2-Ld**2)/(2*Lu*C0))
        A0 = math.sqrt(Au**2+Lw**2-2*Au*Lw*math.cos(agn+agm))
        if A0*1.8<=Data_dict[count]['A']:     #油壓鋼伸長後長度必須小於原長度*1.5
            print("油壓鋼長度不夠")
            count = count+1
            continue
        if Data_dict[count]['A']>=Au+Aw or Au>=Aw+Data_dict[count]['A'] or Aw>=Au+Data_dict[count]['A']:        #A, Au, Aw 圍成一個三角形
            print("2公尺時, 油壓鋼三角形不合")
            count=count+1
            continue

        ag3 = math.pi-(ag1+ag2)
        ag4 = math.acos((Aw**2-Au**2-Data["A"]**2)/(-2*Au*Data["A"]))
        ag5 = ag3-ag4

#受力計算 (驗算完成，Fd, Fe 與C計算出來的有一點點誤差)----------------------------------------------------------------------------------------------------------
        
        Data_dict[count]["Fb"] = Force_/math.sin(agb)
        Data_dict[count]["Fd"] = (Data["Fb"]*math.sin(ag3)*Lu)/(math.sin(ag4)*Au)
        Data_dict[count]["Fe"] = (Data["Fb"]**2+Data["Fd"]**2-2*Data["Fb"]*Data["Fd"]*math.cos(ag5))**0.5
        ag6 = math.asin((Data["Fb"]*math.sin(ag5))/Data["Fe"])
        ag7 = ag4-ag6+age-math.pi/2
        ag8 = ag4-ag6
        Data_dict[count]["ag_Fb"] = ag3
        Data_dict[count]["ag_Fd"] = ag4
        Data_dict[count]["ag_Fe"] = ag8

#PIN (雙截面)應力和尺寸計算 (驗算完成，有改公式，所以跟C輸出不一樣)--------------------------------------------------------------------------------------------
        #剪切力計算結果
        Data_dict[count]["Pin_t"] = ((4*Data["Fb"]*SF)/(Syt*math.pi))**0.5
        #彎矩計算結果
        Data_dict[count]["Pin_b"] = np.cbrt((SF*6*Data["Fb"]*Up_b)/(Syt*math.pi))

        
#Down 下拉桿應力和尺寸計算-----------------------------------------------------------------------------------------------------------------------------------
        i1 = 0
        Out_SF = [None]*30
        for Data_D in Down_Data_dict:
            Down_Data_dict_Temp = Down_Data_dict        #SF暫存區

            Out_SF[i1] = (Syt*2*Data_D["Down_t"]*(Data_D["Down_h"]+Data_D["Down_b"]-2*Data_D["Down_t"]))/Data["Fb"]
            Down_Data_dict[i1]["Down_SF"] = Out_SF[i1]
            #print(f'SF{i1+1} = {Down_Data_dict[i1]["Down_SF"]}')     #把每一個截面尺寸的SF列印出來
            i1 = i1+1
        print(f' {count+1}號桿子尺寸，下拉桿的"所有截面尺寸"的 Down_SF_List = {Out_SF}')     #把以上算好個截面尺寸的SF用List形式全部列印出來

    #找出最接近2的SF的規格, 並存入Data_dict[count]["Down_SF"]、Data_dict[count]["Down_h"]、Data_dict[count]["Down_b"]、Data_dict[count]["Down_t"]
        Out_SF = list(filter(lambda x: x is not None and x >= 0, Out_SF))       #去除掉Out_SF_2中的負數和None

        if Out_SF:        #為防止Out_SF_2中去除負數和None後變成空陣列。如果是空陣列則跳出迴圈()
            Data_dict[count]["Down_SF"] = min(Out_SF, key=lambda x: abs(x - 2))
        else:
            # 在這裡處理列表為空的情況
            print("Out_SF 列表為空")
            count=count+1
            continue

        Data_dict[count]["Down_SF"] = min(Out_SF, key=lambda x: abs(x - 2))         #找出SF最接近2的值，並存入 Data_dict[count]["Down_SF"]
        #print("下拉桿: 選擇截面尺寸的SF(最接近2):", Data_dict[count]["Down_SF"])  
        #找出SF最接近2的值的規格
        index = find_index_in_array(Down_Data_dict_Temp, Data_dict[count]["Down_SF"])      #SF最接近2的值，在第index+1項
        #print(f"在Data_dict內陣列的第 {index+1} 項")
        Data_dict[count]["Down_h"] = Down_Data_dict[index]["Down_h"]
        Data_dict[count]["Down_b"] = Down_Data_dict[index]["Down_b"]
        Data_dict[count]["Down_t"] = Down_Data_dict[index]["Down_t"]
        

#Up 上拉桿應力和尺寸計算------------------------------------------------------------------------------------------------------------------------------------
        i2 = 0
        Out_SF_2 = [None]*40        #截面尺寸算完的SF會存入Out_SF_2，最多可以算30個截面尺寸
        for Data_U in Up_Data_dict:
    #公式區
            Mmax = Data["Fb"]*math.sin(ag3)*(Lu-Au)
            Axial_Load = (Data["Fb"]*math.cos(ag3))/(2*Data_U["Up_t"]*(Data_U["Up_h"]+Data_U["Up_b"]-2*Data_U["Up_t"]))
            I_RecTube = ((Data_U["Up_b"]*Data_U["Up_h"]**3-(Data_U["Up_b"]-2*Data_U["Up_t"])*(Data_U["Up_h"]-2*Data_U["Up_t"])**3)/12)
            Out_SF_2[i2] = (Syt*I_RecTube*2)/(Mmax*Data_U["Up_h"]+I_RecTube*Axial_Load)
            Up_Data_dict[i2]["Up_SF"] = Out_SF_2[i2]
            #print(f'Up_SF{i2+1} = {Up_Data_dict[i2]["Up_SF"]}')     #把每一個截面尺寸的SF列印出來
            i2 = i2+1

        print(f' {count+1}號桿子尺寸，上拉桿的"所有截面尺寸"的 Up_SF_List  = {Out_SF_2}')     #把以上算好個截面尺寸的SF用List形式全部列印出來

    #找出最接近2的SF的規格，並存入Data_dict[count]["Up_SF"]、Data_dict[count]["Up_h"]、Data_dict[count]["Up_b"]、Data_dict[count]["Up_t"]
        Out_SF_2 = list(filter(lambda x: x is not None and x >= 0, Out_SF_2))       #去除掉Out_SF_2中的負數和None
        
        if Out_SF_2:        #為防止Out_SF_2中去除負數和None後變成空陣列。如果是空陣列則跳出迴圈()
            Data_dict[count]["Up_SF"] = min(Out_SF_2, key=lambda x: abs(x - 2))
        else:
            # 在這裡處理列表為空的情況
            print("Out_SF_2 列表為空")
            count=count+1
            continue
        

        Data_dict[count]["Up_SF"] = min((x for x in Out_SF_2 if x > 2), default=None, key=lambda x: abs(x - 2))        #找出SF最接近2的值，並存入 Data_dict[count]["Up_SF"]
        #print("上拉桿: 選擇截面尺寸的SF(最接近2)", Data_dict[count]["Up_SF"])        
        #找出SF最接近2的值的規格
        
        index2 = find_index_in_array(Up_Data_dict, Data_dict[count]["Up_SF"])      #SF最接近2的值，在第index2+1項
        '''
        if Data_dict[count]["Up_SF"] <2:
            Data_dict[count]["Up_SF"] = 2
            index2+=1
        '''
        #print(f"在Data_dict內陣列的第 {index2+1} 項")
        Data_dict[count]["Up_h"] = Up_Data_dict[index2]["Up_h"]
        Data_dict[count]["Up_b"] = Up_Data_dict[index2]["Up_b"]
        Data_dict[count]["Up_t"] = Up_Data_dict[index2]["Up_t"]

        count=count+1



    output_df = pd.DataFrame(Data_dict)     #"陣列中字典"形式 轉換 DataFrame
    output_df.to_excel("MD_2_Output.xlsx")
    print(output_df)
    print("輸出完成，程式結束")


