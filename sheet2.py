import openpyxl
import shutil
import os 
import xlwings as xw

def AramcoSheet2 (csvpath): 
          
        '''
    Read Aramco cvs data: P&RMU NGL Spheres D-011 A/B Level Calcuations  
    and take all function needed for the calulation to process it with the user input
    
    Attributes:
    -----------
        csvpath: path to the csv file 
    '''
        
#         # create a copy of existing csv
#         dirname = os.path.dirname(csvpath)
#         basename = os.path.basename(csvpath) # get the filename
        
#         targetpath = dirname + "/copy"+ basename
#         #print (targetpath)
        
    
        
          # read the excel sheet 
        wb2 = openpyxl.load_workbook(csvpath)
        ws2 = wb2['NGL Sphere']
        
        
        
        historainInputCase = {"Initial Level (%)": "E8", "Final Level (%)": "E10"}
        historainOutput = {"final volume (bbl)": "E11", 
                          "Time (h)" : "E12", 
                          "Time (h)  [2 spheres]": "E13"}
        
    
        manualInputcase = {"Total Flow from LRU (MBD)":"E16",
                   "Total flow from Inlet (MBD)":"E17",
                  "Total NGL production (MBD)":"E18", 
                  "Initial Level (%)": "E20",
                  "Final Level (%)": "E22"}
        manualOutput = {"final volume (bbl)": "E23", 
                          "Time (h)" : "E24", 
                          "Time (h)  [2 spheres]": "E25"}
        
        Output = {}

        
        # take the input from the user 
    
        userinput= input("Do you want to use the historain values for caculation? Y/N")
        
        if userinput.lower() == "y":
            
            #read the needed input from the user 
            
            for userinput in historainInputCase: 
                hisIn= input(userinput)
                ws2[historainInputCase[userinput]]=float(hisIn)/100
           
            # save the orignal modifications to the csv    
            wb2.save(csvpath)
            
            # read the values of the excel sheet 

            wbxl=xw.Book(csvpath)
            
            for useroutput in historainOutput: 
                Output[useroutput]=wbxl.sheets['NGL Sphere'].range(historainOutput[useroutput]).value

                
                
        elif userinput.lower() == "n":
  
            # read the needed input from the user
            count = 0 
            for userinput in manualInputcase: 
                mIn= input(userinput)
        
                # condition for Percentage cells 
                if count ==4 or count == 6 : 
                    ws2[manualInputcase[userinput]]=float(mIn)/100
                else: 
                    ws2[manualInputcase[userinput]]=float(mIn)
            
             # save the orignal modifications to the csv    
            wb2.save(csvpath)
            
              # read the values of the excel sheet 

            wbxl=xw.Book(csvpath)
            
            for useroutput in manualOutput: 
                Output[useroutput]=wbxl.sheets['NGL Sphere'].range(historainOutput[useroutput]).value

                    
            
            
        #print(Output)
        return Output

            
        
        
test= AramcoSheet2("NGL Sphere Calculations (D-011 AB)_2.xlsx")