#encoding: utf-8
require 'roo'

# ////////////////////////////////////////////////////////////////////////////////
# /// Please rewrite the following variables /////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////

# enter the targe table name.
tblName = "sample_table"

# enter the tatget table colmun name.
columnSize = 19

# ////////////////////////////////////////////////////////////////////////////////
# /// Please rewrite the above variables /////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////






# define column offset / * And SQL comment for sample table.
columnOffset = 2
rowOffset = 4


lastColumnNumber = columnSize + columnOffset

# loading xlsx data.
xlsx = Roo::Excelx.new('./in/' + tblName + '_data.xlsx')

# sql file generation.
sqlFile = File.open('ins_' + tblName + '_base.sql',"w")


# sql execution start statement.
sqlFile.puts("TRUNCATE TABLE " + tblName + "_base;")
sqlFile.puts()
sqlFile.puts("START TRANSACTION;")
sqlFile.puts()
sqlFile.puts("INSERT INTO " + tblName + "_base VALUES")



rowIndex = 0;
xlsx.each_row_streaming(offset: rowOffset) do |row|

	# define variables.
    shop_type = row[2].nil? || row[2].empty? ? 0 : row[2]
    banner_id = row[3].nil? || row[3].empty? ? 0 : row[3]
    sort_id = row[4].nil? || row[4].empty? ? 0 : row[4]

    startDateTime_y = row[5].nil? || row[5].empty? ? 0 : row[5].value
    startDateTime_m = row[6].nil? || row[6].empty? ? 0 : row[6].value
    startDateTime_d = row[7].nil? || row[7].empty? ? 0 : row[7].value
    startDateTime_H = row[8].nil? || row[8].empty? ? 0 : row[8].value
    startDateTime_M = row[9].nil? || row[9].empty? ? 0 : row[9].value
    startDateTime_S = row[10].nil? || row[10].empty? ? 0 : row[10].value

    endDateTime_y = row[11].nil? || row[11].empty? ? 0 : row[11].value
    endDateTime_m = row[12].nil? || row[12].empty? ? 0 : row[12].value
    endDateTime_d = row[13].nil? || row[13].empty? ? 0 : row[13].value
    endDateTime_H = row[14].nil? || row[14].empty? ? 0 : row[14].value
    endDateTime_M = row[15].nil? || row[15].empty? ? 0 : row[15].value
    endDateTime_S = row[16].nil? || row[16].empty? ? 0 : row[16].value

    startDateTime = sprintf("%4d-%02d-%02d %2d:%02d:%02d", startDateTime_y, startDateTime_m, startDateTime_d, startDateTime_H, startDateTime_M, startDateTime_S)
	endDateTime = sprintf("%4d-%02d-%02d %2d:%02d:%02d", endDateTime_y, endDateTime_m, endDateTime_d, endDateTime_H, endDateTime_M, endDateTime_S)

	image_path = row[17].nil? || row[17].empty? ? "" : "#{row[17]}"
	item_category = row[18].nil? || row[18].empty? ? 0 : row[18]
	item_id = row[19].nil? || row[19].empty? ? 0 : row[19]


	insertParamArray = []
	insertParamArray.push(shop_type, banner_id, sort_id, '"' + startDateTime + '"', '"' + endDateTime + '"', '"' + image_path.to_s + '"', item_category, item_id)


    if rowIndex == 0
        sqlFile.print(" ( ")
    else
        sqlFile.print(",( ")
    end

	sqlFile.print( insertParamArray.join(', ') )
    sqlFile.print(" )") 

	sqlFile.puts()
    rowIndex += 1
end

sqlFile.puts(";")
sqlFile.puts()
sqlFile.puts("COMMIT;")




