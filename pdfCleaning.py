from camelot import read_pdf
import pandas as pd



def pdf_extraction(pdfFileName):
#since it works on table_areas function if empty pages are provided it will cause error, hence remove empty page
    my_name=read_pdf(pdfFileName,pages="1",flavor='stream', table_areas=['64,663,121,22'],row_tol=30)
    my_mobile=read_pdf(pdfFileName,pages="1",flavor='stream', table_areas=['173,661,233,37'],strip_text=".\n",row_tol=30)
    my_address=read_pdf(pdfFileName,pages="1",flavor='stream', table_areas=['431,665,529,24'],row_tol=40)
    
    
    my_name._tables.extend(read_pdf(pdfFileName,pages="2-end",flavor='stream', table_areas=['60,771,121,22'],row_tol=30))
    my_mobile._tables.extend(read_pdf(pdfFileName,pages="2-end",flavor='stream', table_areas=['170,770,233,36'],strip_text=".\n",row_tol=30))
    my_address._tables.extend(read_pdf(pdfFileName,pages="2-end",flavor='stream', table_areas=['430,773,665,24'],row_tol=40))
    
    file_name="DLUpdatedList.xlsx"
    printing_sheet=pd.DataFrame(columns=["Name","Registration","Address"],data=None)
    print(len(my_name),len(my_mobile),len(my_address))
    
    user=0
    for index in range(len(my_name)):
        user_name=my_name[index].df[0]
        user_mobile=my_mobile[index].df[0]
        user_address=my_address[index].df[0] 
        for i in range(len(user_name)):
            user=user+1
            #print(user)
            #print(user_name[i],user_mobile[i],user_address[i],"\n")
            printing_sheet.at[user,"Name"]=user_name[i]
            printing_sheet.at[user,"Registration"]=user_mobile[i]
            printing_sheet.at[user,"Address"]=user_address[i]
        
    
    printing_sheet.to_excel(file_name)
    
    return 1
    
"""
if __name__=="__main__":
    pdf_extraction("/Users/JaySabnis/Downloads/DL_REPORT.pdf")
"""