import pandas as pd
import numpy as np
import math

if __name__ == '__main__':

    Length_dict = []

    #輸出所有長度尺寸

    B = [2750]        #假設條件1  #1.間隔(無)
    LW = [1500,2000]    #假設條件2  #2.間隔
    for b in B:
        print("1. B= ",B)

        for lw in LW:
            AW =lw+100      #假設條件3: LW<=AU<=2000+LW
            LU = 1200       #假設條件4: 1200<=LU<=3000
            print(" 2. LW= ",lw)

            while AW <= 2000+lw:
                print("  3. AW= ",AW)

                while LU <= 3000:
                    print("   4. LU= ",LU)
                    LD = 500        #假設條件5: 500<=LD<=3000

                    
                    while LD<= 3000:
                        if LD <= (b**2+(lw+2000)**2)**0.5+LU and (b**2+(lw+2000)**2)**0.5 <= LD+LU and LU <= (b**2+(lw+2000)**2)**0.5+LD:       #必要條件
                            print("   5. LD= ",LD)
                            AU = 1000       #假設條件6: 1000<=AU<=LU-500

                            while AU <= LU-500:

                                Length_dict.append({"Base_L": b, "Wall_L": lw, "Down_L": LD, "Up_L": LU, "Wall_A": AW, "Up_A": AU, 'A': None, 'Fb': None, 'ag_Fb': None, 'Fd': None, 'ag_Fd': None, 'Fe': None, 'ag_Fe': None, 'Pin_t': None, 'Pin_b': None, 'Down_SF': None, 'Down_h': None, 'Down_b': None, 'Down_t': None, 'Up_SF': None, 'Up_h': None, 'Up_b': None, 'Up_t': None})
                                print("     6. AU= ",AU)

                                AU+=100    #6.間隔
                        LD+= 100    #5.間隔

                    LU+= 50    #4.間隔

                AW+= 100    #3.間隔

    Length_DataFarme = pd.DataFrame(Length_dict)
    print(Length_DataFarme)
    Length_DataFarme.to_excel("Length_DataFarme.xlsx")

    print("編輯成功，程式結束")
