#-*- coding: utf-8 -*-
require "java"
require "lib/poi-3.8-20120326.jar"
require "lib/poi-ooxml-3.8-20120326.jar"

java_import java.io.FileInputStream
java_import java.io.FileOutputStream
java_import org.apache.poi.ss.usermodel.WorkbookFactory
java_import org.apache.poi.ss.usermodel.DataFormatter

class ExcelWorkbook
  def initialize(path)
    @fis = FileInputStream.new(path)
    @book = WorkbookFactory::create(@fis)
  end

  def self.open(path, &block)
    book = ExcelWorkbook.new(path)
    return book unless block_given?
    yield book
  ensure
    book.close
  end

  def write(path)
    fos = FileOutputStream.new(path)
    @book.write(fos)
  ensure
    fos.close
  end

  def close
    @fis.close
  end

  def select_sheet_at(index)
    @sheet = @book.get_sheet_at(index)
    @sheet.set_force_formula_recalculation(true)
  end

  def []=(row_idx, col_idx, value)
    @sheet.create_row(row_idx) if @sheet.get_row(row_idx).nil?
    row = @sheet.get_row(row_idx)
    row.create_cell(col_idx) if row.get_cell(col_idx).nil?
    row.get_cell(col_idx).set_cell_value(value)
  end

  def [](row, col)
    return nil if (row = @sheet.get_row(row)).nil?
    return nil if (cell = row.get_cell(col)).nil?

    return DataFormatter.new.format_cell_value(cell)
  end

  def to_blob
    # use tempfile class
    tmp_io = Tempfile.open("excel_server_creating_excel_file_done_")
    write(tmp_io.path)
    
    blob = nil
    tfp = open(tmp_io.path, "r+b")
    blob = tfp.read
    tfp.close
    
    tmp_io.close(true)
    return blob
  end
end

=begin
# ----
# sample
if __FILE__ == $0
  excel_filename = "template_excel_file.xls"
  filepath = File.expand_path("./", excel_filename)
  ExcelWorkbook.open(filepath) do |book|
  #book = ExcelWorkbook.new(filepath)
    book.select_sheet_at(0)
    book[0, 0] = "hoge"
    book[1, 9] = "日本語おk？"
    book.write("new_excel.xls")
  #book.close
  end
end
=end
