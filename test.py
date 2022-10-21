# import xlsxwriter module
import xlsxwriter

# which is the filename that we want to create.
#create a workbook
workbook = xlsxwriter.Workbook("example.xlsx")

# The workbook object is then used to add new worksheet via the add_worksheet() method.
#add a worksheet
worksheet = workbook.add_worksheet()

# Create a format to use in the merged range
merge_format = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#C0C0C0', #hex format for silver colour
    'border': 1,
    'text_wrap': True})

worksheet.merge_range('D5:E5', 'Test Details', merge_format)
worksheet.merge_range('A5:A8', 'Heap Size', merge_format)
worksheet.merge_range('B5:B8', 'Property Settings', merge_format)
worksheet.merge_range('C5:C8', 'Test Rounds', merge_format)
worksheet.merge_range('D6:D8', 'Concurrent Users', merge_format)
worksheet.merge_range('E6:E8', 'Test Duration(min)', merge_format)
worksheet.merge_range('G6:I6', 'Response time(ms)', merge_format)
worksheet.merge_range('F5:K5', 'JMeter Outputs', merge_format)
worksheet.merge_range('F6:F8', 'Total Requests', merge_format)
worksheet.merge_range('L6:M6', 'BPM JVM Stats', merge_format)
worksheet.merge_range('N6:O6', 'BPM Docker Stats', merge_format)
worksheet.merge_range('L7:L8', 'Avg.CPUUsage(%)', merge_format)
worksheet.merge_range('M7:M8', 'Avg.MemUsage(MB)', merge_format)
worksheet.merge_range('N7:N8', 'Avg.CPUUsage(%)', merge_format)
worksheet.merge_range('O7:O8', 'Avg.MemUsage(MB)', merge_format)
worksheet.merge_range('G7:G8', 'Average', merge_format)
worksheet.merge_range('H7:H8', 'Min', merge_format)
worksheet.merge_range('I7:I8', 'Max', merge_format)
worksheet.merge_range('J7:J8', 'Error', merge_format)
worksheet.merge_range('K7:K8', 'Throughput', merge_format)
worksheet.merge_range('J6:K6', '', merge_format)
worksheet.merge_range('L5:O5', '', merge_format)

worksheet.merge_range('P6:P8', 'Process Instance Started /second', merge_format)
worksheet.merge_range('Q6:Q8', 'Process Instance Completed /second', merge_format)
worksheet.merge_range('R6:R8', 'Tasks Started /second', merge_format)
worksheet.merge_range('S6:S8', 'Tasks Completed /second', merge_format)
worksheet.merge_range('T6:T8', 'Sub-procs started /second', merge_format)
worksheet.merge_range('U5:U8', 'Total Execution Time', merge_format)
worksheet.merge_range('P5:T5', 'Database Stats', merge_format)
#worksheet.merge_range('O7:P7', 'Avg.MemUsage(MB)', merge_format)

#set column width
worksheet.set_column('A:B', 8)
worksheet.set_column('C:D', 14)
worksheet.set_column('F:K', 14)
worksheet.set_column('G:H', 8)
worksheet.set_column('I:J', 8)
worksheet.set_column('E:L', 15)
worksheet.set_column('M:N', 15)
worksheet.set_column('O:P', 15)
worksheet.set_column('Q:R', 10)
worksheet.set_column('S:T', 10)
worksheet.set_column('U:V', 10)
#worksheet.set_column('U:V', 15)

# Create a format to use in the merged range
merge_format2 = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#e6e6fa'})#hex format for levender colour
#merge range of cells over rows and columns
worksheet.merge_range('A3:U4', 'BPM530_V159.2 + DB2', merge_format2)
merge_format2.set_font_size(20)


#add borders to the worksheet
#conditional_format -apply formatting based on user defined criteria
border_format=workbook.add_format({'border':1})

worksheet.conditional_format( 'A9:U11' , { 'type' : 'blanks' , 'format' : border_format} )
worksheet.conditional_format( 'A9:U11' , { 'type' : 'no_blanks' , 'format' : border_format} )


# Finally, close the Excel file via the close() method.
workbook.close()
