require 'rubygems'
require 'stringio'
require 'simple_xlsx'
require 'tempfile'
require 'tmpdir'

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
      path = File.join(Dir.tmpdir, "serializer-#{rand}")
      write_file(path)

      File.open(path) do |file|
        io.write(file.read)
      end

      File.unlink(path)

      io
    end

    def write_file(path)
      puts path

      SimpleXlsx::Serializer.new(path) do |book|
        write_description(book) if @options[:description]
        write_book(book)
      end

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
      case model
      when Array
        sheet.add_row(model)
      else
        sheet.add_row(columns.collect { |key|
          model.send(key)
        })
      end
    end
  end
end
