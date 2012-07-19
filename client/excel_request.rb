#-*- coding: utf-8 -*-
# -----------------------------------------------------------------------------
# excel create request
#
# code sample 
#
# ereq = ExcelRequest.new
# templatefilepath = File.expand_path("./", "template_excel_file.xls")
#
# sheets = ExcelRequest::ExcelSheet.new
# sheet = sheets[0]
# sheet[0, 0] = "hoge"
# sheet[1, 9] = "日本語ですーどうですか？"
# sheet[4, 8] = "ほがー"
#
# write_filepath = File.expand_path("./", "newfile.xls")
# ereq.get(templatefilepath, sheets, write_filepath)
#
# => create 'newfile.xls' using ExcelServer using template_file
# 
# -----------------------------------------------------------------------------

require "rubygems"
require "net/http"
require "uri"
require "cgi"
require "json"
require "base64"

class ExcelRequest
  class ExcelSheet
    def initialize
      @sheets = Hash.new
    end
    def [](index)
      return @sheets[index] if @sheets.key?(index)
      @sheets[index] = ExcelCell.new
    end
    def to_json
      newhash = {}
      @sheets.each do |idx, val|
        newhash[idx] = Array.new
        val.cell.each do |k, v|
          newhash[idx].push([k[0], k[1], v])
        end
      end
      newhash.to_json
    end
  end
  
  class ExcelCell
    attr_reader :cell
    def initialize
      @cell = Hash.new
    end
    def []=(row, col, value)
      @cell[[row, col]] = value
    end
    def [](row, col)
      @cell[[row, col]]
    end
  end
  
  def initialize(excelserver = "http://localhost:9292")
    @excelserver = excelserver
    @sheets = nil

    @template_filepath = nil
  end

  def set_exceltemplate(filepath)
    @template_filepath = filepath
  end
  
  def set_exceldata(excelsheets)
    if excelsheets.instance_of?(ExcelRequest::ExcelSheet)
      @sheets = excelsheets
    else
      raise
    end
  end

  def get(template_filepath, excelsheets, new_filepath)
    set_exceltemplate(template_filepath)
    set_exceldata(excelsheets)

    response = send_request
    return false if response == false

    tfp = open(new_filepath, "w+b")
    tfp.write Base64.decode64(response)
    tfp.close
    return true
  end

  def get_blob(template_filepath, excelsheets)
    set_exceltemplate(template_filepath)
    set_exceldata(excelsheets)

    response = send_request
    return nil if response == false

    return Base64.decode64(response)
  end

  private

  def send_request(url = "/excel_server/create_excelfile")
    return false if @template_filepath.nil?
    return false if @sheets.nil?
    
    encoded_template = get_exceltemplate(@template_filepath)
    datas = @sheets.to_json

    params = {:template => encoded_template, :datas => datas}
    response = request(url, params)
    
    return response
  end
  
  def get_exceltemplate(filepath = @template_filepath)
    return false if filepath.nil?

    blob = nil
    tfp = open(filepath, "r+b")
    blob = tfp.read
    tfp.close
    
    return Base64.encode64(blob)
  end

  def request(url, params = {})
    requestaddress = @excelserver + url

    default_params = {:newfile => nil}
    request_params = default_params.merge(params)

    uri = URI.parse(requestaddress)
    response = nil
    Net::HTTP.start(uri.host, uri.port){|http|
      request = Net::HTTP::Post.new(uri.path)
      request.set_form_data(request_params)
      response = http.request(request)
    }
    return response.body.strip
  end
end
