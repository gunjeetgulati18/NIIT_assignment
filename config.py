'''
config provided to set all the variables needed in script
'''


input1_url = 'http://www.bcb.gov.br/pec/Indeco/Ingl/ie5-24i.xlsx'   # needed to download the file
input2_url = 'http://www.bcb.gov.br/pec/Indeco/Ingl/ie5-26i.xlsx'	# needed to download the file
input1_filename = 'input1'
input2_filename = 'input2'
output1_filename = 'bcb_output_1.csv'							
output2_filename = 'bcb_output_2.csv'	
json_output1_filename = 'bcb_output_1.txt'  			 #  file type can be .txt or .json 							
json_output2_filename = 'bcb_output_2.txt'				#  file type can be .txt or .json 							
output1_ref = 'ref1.csv'								# needed to validate the previous records
output2_ref = 'ref2.csv'								# needed to validate the previous records
input1_header = [['Date','BCB_Commercial_Exports_Total','BCB_Commercial_Exports_Advances_on_Contracts','BCB_Commercial_Exports_Payment_Advance','BCB_Commercial_Exports_Others','BCB_Commercial_Imports','BCB_Commercial_Balance','BCB_Financial_Purchases','BCB_Financial_Sales','BCB_Financial_Balance','BCB_Balance']]
input2_header = [['Date','BCB_FX_Position']]
