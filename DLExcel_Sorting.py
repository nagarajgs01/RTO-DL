

import pandas as pd
import ctypes


def dlExcel_sort():
    
     fileName="DLUpdatedList.xlsx"
     df=pd.read_excel(fileName,header=0,index_col=0,dtype={"Address":str})
     
    
     applicant=0
     
     df["Pin"]=None
     df["Mobile"]=None
     
     
     for index in df["Registration"]:
         
         applicant=applicant+1
         details=df[df["Registration"]==index]
       
        
         address=details.iloc[:,2].values.tolist()
        # print(address)
         list_variable=address[0].split("\n")
        # print(list_variable)
         father_name=list_variable.pop(0)
        # print(list_variable)
         
         address=" ".join(list_variable)
         pinaddress="".join(list_variable)
         #print(pinaddress)
         
         
         address=str(address)
         
         pincode=pinaddress[-6:]
         #print(pincode)
         address=address.strip(pincode)
         #print(address)
         df.at[applicant,"Address"]=address
         df.at[applicant,"Pin"]=pincode
        
            
            
        #spliting Dl and mobile number
         reg=details.iloc[:,1].values.tolist()
         reg=reg.pop()
         #print(reg,len(reg))
         
         mobileNumber=reg[-10:]
         
         #print(mobileNumber,len(mobileNumber))
        
         reg=reg[0:15]
         #print(reg,len(reg))
         
         df.at[applicant,"Registration"]=reg
         df.at[applicant,"Mobile"]=mobileNumber
         
         
         #modifying name
         name=details["Name"].values.tolist()
         
         name=name.pop()
         
         name=str.join(" ",name.splitlines())
         #print(name)
         name=name.split(" ")
         name.pop()
         name=" ".join(name)
        # print(name)
         
    
         df.at[applicant,"Name"]=name
     
         print(name,reg,mobileNumber,address,pincode,"\n")
             
     df.to_excel(fileName)
        
     print("file saved \n")
    
         
def print_card(values):
    
  
    file_name="Dlprint_details.xlsx"
    
    new_excel=pd.read_excel("DLUpdatedList.xlsx",header=0,index_col=0,dtype={"Number":str})
    
    printing_sheet=pd.DataFrame(columns=["Registration","Name","Address","Pin","Mobile"],data=None)
    printing_sheet.to_excel(file_name)
    
    
    #print(new_excel.head())
    printing_sheet=pd.read_excel(file_name,header=0,index_col=0,dtype={"Number":str})
    #print(printing_sheet.head())
    user=0
    for index in values: 
        user=user+1
        
        val=user_information=new_excel[new_excel["Registration"]==index]
        
        if len(val):
 
            user_information=new_excel[new_excel["Registration"]==index]
            
            temp1=str(user_information['Registration'].values.tolist()).strip("[]").strip("'")
            printing_sheet.at[user,"Registration"]=temp1
            
            temp1=str(user_information['Name'].values.tolist()).strip("[]").strip("'")
            printing_sheet.at[user,"Name"]=temp1
            
            temp1=str(user_information['Address'].values.tolist()).strip("[]").strip("'")
            printing_sheet.at[user,"Address"]=temp1 
            
            temp1=str(user_information['Mobile'].values.tolist()).strip("[]").strip("'")
            printing_sheet.at[user,"Mobile"]=temp1 
            
            temp1=str(user_information['Pin'].values.tolist()).strip("[]").strip("'")
            printing_sheet.at[user,"Pin"]=temp1 
        
      
       
       
    
    
    
    printing_sheet.to_excel(file_name)
    
         
    
    
    
    
"""
if __name__=="__main__":
   # dlExcel_sort("DLUpdatedList.xlsx")
   values=[]
   print_card(values)
"""