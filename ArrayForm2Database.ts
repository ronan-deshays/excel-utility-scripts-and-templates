function main(workbook: ExcelScript.Workbook) {
    
    // define a few names
      // worksheets
    const PARAM_SHEET_NAME = "Parametres"
    const SCHEDULE_SHEET_NAME = "Planning"
    const DATABASE_SHEET_NAME = "BaseDeDonnees"
      // tables
    const DATABASE_TABLE_NAME = "Database"
      // ranges
    const NUMBER_OF_COLUMNS_RANGE_NAME = "nb_col"
    const NUMBER_OF_ROWS_RANGE_NAME = "nb_row"
    const FIRST_COL_NUMBER_RANGE_NAME = "num_first_col"
    const FIRST_ROW_NUMBER_RANGE_NAME = "num_first_row"
      // internal script param
    const AFFAIR_NUM_COL_INDEX = 0 // zero based index in array of values from a range given by user
    const AFFAIR_NAME_COL_INDEX = 1
    const AFFAIR_DATE_ROW_INDEX = 0
    const AFFAIR_DAY_TIME_ROW_INDEX = 1
  
    // init variables
    let param_sheet = workbook.getWorksheet(PARAM_SHEET_NAME)
    let schedule_sheet = workbook.getWorksheet(SCHEDULE_SHEET_NAME)
    let database_sheet = workbook.getWorksheet(DATABASE_SHEET_NAME)
    let database_table = workbook.getTable(DATABASE_TABLE_NAME)
  
    // clear table (it is a full update)
    if (database_table.getRowCount() > 0) { // if database not already empty
      database_table.deleteRowsAt(0, database_table.getRowCount()) // delete all rows in database
    }
  
    // get user param
    let nb_row = param_sheet.getRange(NUMBER_OF_ROWS_RANGE_NAME).getValue() as number
    let nb_col = param_sheet.getRange(NUMBER_OF_COLUMNS_RANGE_NAME).getValue() as number
    let num_first_col = param_sheet.getRange(FIRST_COL_NUMBER_RANGE_NAME).getValue() as number - 1 // convert from 1 indexed to 0 indexed
    let num_first_row = param_sheet.getRange(FIRST_ROW_NUMBER_RANGE_NAME).getValue() as number - 1
  
    // get data (data structure is as follow : [Excel_row][Excel_col])
    let schedule_data = schedule_sheet.getRangeByIndexes(num_first_row, num_first_col, nb_row, nb_col).getValues()
    let dates_data = schedule_sheet.getRangeByIndexes(0,num_first_col,2,nb_col).getValues()
    let affairs_data = schedule_sheet.getRangeByIndexes(num_first_row, 0, nb_row, 2).getValues()
  
    // loop through source range
    for (let row_index = 0; row_index < (nb_row - 1); row_index++)
    {
      for (let col_index = 0; col_index < (nb_col - 1); col_index++) 
      {
        if (schedule_data[row_index][col_index] != "")
        {
          // save reorganised data to table
          database_table.addRow(
            0,
            [
              affairs_data[row_index][AFFAIR_NUM_COL_INDEX],
              affairs_data[row_index][AFFAIR_NAME_COL_INDEX],
              dates_data[AFFAIR_DATE_ROW_INDEX][col_index],
              dates_data[AFFAIR_DAY_TIME_ROW_INDEX][col_index],
              schedule_data[row_index][col_index]
            ]
          )
        }
      }
    }
  }
  