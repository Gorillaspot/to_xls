require 'rubygems'
require 'stringio'
require 'spreadsheet'

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
      book = Spreadsheet::Workbook.new
      write_description(book) if @options[:description]
      write_book(book)
      book.write(io)
    end

    def write_description(book)
      sheet = book.create_worksheet
      sheet.name = "Description"

      row_index = 0
      row = sheet.row(row_index)
      case @options[:description]
      when String, Symbol
        row.push(option[:description])
      when Array
        @options[:description].each do |line|
          row = sheet.row(row_index)
          row.push(line)
          row_index += 1
        end
      end
    end

    def write_book(book)
      sheet = book.create_worksheet
      sheet.name = @options[:name] || 'Data'
      write_sheet(sheet)
      return book
    end

    def write_sheet(sheet)
      if columns.any?
        row_index = 0

        if headers_should_be_included?
          fill_row(sheet.row(0), headers)
          row_index = 1
        end

        @array.each do |model|
          row = sheet.row(row_index)
          fill_row(row, columns, model)
          row_index += 1
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

    def fill_row(row, column, model=nil)
      case column
      when String, Symbol
        row.push(model ? model.send(column) : column)
      when Hash
        column.each{|key, values| fill_row(row, values, model && model.send(key))}
      when Array
        column.each{|value| fill_row(row, value, model)}
      else
        raise ArgumentError, "column #{column} has an invalid class (#{ column.class })"
      end
    end

  end

end
