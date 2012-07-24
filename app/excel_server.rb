#-*- coding: utf-8 -*-
class ExcelServer
  namespace "/excel_server" do

    get '/nowtime' do
      h = {:nowtime => Time.now }
      h.to_json
    end
    
    post '/nowtime' do
      h = {:nowtime => Time.now }
      h.to_json    
    end
    
    post '/create_excelfile' do
      template = params[:template]
      datas = params[:datas]
      
      # decoding template
      tmp_blob = Base64.decode64(template)
      
      # excel data
      exceldatas = JSON.parse(datas)
      
      # create tempolary file
      tmp_io = Tempfile.open("excel_server_creating_excel_template_file_")
      tmp_io.write(tmp_blob)
      tmp_filepath = tmp_io.path
      
      blob = "Hello"
      ExcelWorkbook.open(tmp_filepath) do |book|
        exceldatas.each do |idx, val|
          case idx.to_i
          when -1
            val.each do |data|
              cmd, params = data
              cmd.gsub!("@", "")
              begin
                book.send(cmd, *params)
              rescue
                p $!
              end
            end
          else
            book.select_sheet_at(idx.to_i)
            val.each do |data|
              row, col, val = data
              case val.class.to_s
              when "Array" # -> ["value", ["@command", ["params"]], ...]
                v = val.shift
                val.each do |cmds|
                  cmd, params = cmds
                  cmd.gsub!("@", "")
                  params.unshift(col)
                  params.unshift(row)
                  begin
                    book.send(cmd, *params)
                  rescue
                    p $!
                  end
                end
                book[row, col] = v              
              else
                book[row, col] = val
              end
            end
          end
        end
        blob = book.to_blob
      end
      
      tmp_io.close(true)
      return Base64.encode64(blob)
    end
    
  end
end

