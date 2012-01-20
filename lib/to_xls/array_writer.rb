require 'rubygems'
require 'stringio'
require 'simple_xlsx'
require 'tempfile'

module ToXls
  class ArrayWriter
    def initialize(array, options = {})
      @array = array
      @options = options
    end

    def write_string(string = '')
      io = StringIO.new(string)
      write_io(io)
      io.string
    end

    def write_io(io)
      output = Tempfile.new "serializer" 
      book = SimpleXlsx::Serializer.new(output.path)
      write_description(book) if @options[:description]
      write_book(book)
      data = output.read
      output.close 

      io.write(data)
      data
    end

    def write_description(book)
      book.add_sheet("Description") do |sheet|
        case @options[:description]
        when String, Symbol
          sheet.add_row(@options[:description])
        when Array
          @options[:description].each do |line|
            sheet.add_row(line)
          end
        end
      end
    end

    def write_book(book)
      book.add_sheet(@options[:name] || 'Data') do |sheet|
        write_sheet(sheet)
      end
    end

    def write_sheet(sheet)
      if columns.any?
        if headers_should_be_included?
          sheet.add_row(headers)
        end

        @array.each do |model|
          fill_row(sheet, columns, model)
        end
      end
    end

    def columns
      return  @columns if @columns
      @columns = @options[:columns]
      raise ArgumentError.new(":columns (#{columns}) must be an array or nil") unless (@columns.nil? || @columns.is_a?(Array))
      @columns ||=  can_get_columns_from_first_element? ? get_columns_from_first_element : []
    end

    def can_get_columns_from_first_element?
      @array.first && 
      @array.first.respond_to?(:attributes) &&
      @array.first.attributes.respond_to?(:keys) &&
      @array.first.attributes.keys.is_a?(Array)
    end

    def get_columns_from_first_element
      @array.first.attributes.keys.sort_by {|sym| sym.to_s}.collect.to_a
    end

    def headers
      return  @headers if @headers
      @headers = @options[:headers] || columns
      raise ArgumentError, ":headers (#{@headers.inspect}) must be an array" unless @headers.is_a? Array
      @headers
    end

    def headers_should_be_included?
      @options[:headers] != false
    end

private
    def fill_row(sheet, columns, model=nil)
      case columns
      when Hash
        sheet.add_row(columns.keys.collect { |key|
          if model and v=model.send(key)
              v
          else
            columns.keys[key]
          end
        })
      when Array
        sheet.add_row(columns.collect { |column|
          model ? model.send(column) : column
        })
      else
        raise ArgumentError, "column #{column} has an invalid class (#{ column.class })"
      end
    end
  end
end
