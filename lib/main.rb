require 'spreadsheet'

class Main

  workbook = Spreadsheet.open 'lib/files/source.xls'
  new_book = Spreadsheet::Workbook.new

  new_book.create_worksheet name: 'output'

  format = Spreadsheet::Format.new weight: :bold

  workbook.worksheets.each do |sheet|
    table_name = sheet.row(0)[4]

    sheet.each do |row|
      if row[4] != table_name
        header = ['Campo', 'Tipo', 'Descrição', 'Observação']
        new_book.worksheet(0).insert_row(0, header)
        new_book.worksheet(0).insert_row(0, [table_name])
        new_book.worksheet(0).insert_row(0, [row[0], row[1], row[2], row[3]])

        row.set_format(0, format)

        table_name = row[4]
      else
        new_book.worksheet(0).insert_row(0, [row[0], row[1], row[2], row[3]])
      end
    end
  end

  new_book.write('output.xls')
end
