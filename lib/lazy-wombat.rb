# encoding: utf-8
require 'lazy-wombat/version'
require 'axlsx'

module LazyWombat

  class Xlsx

    def initialize options = {}, &block
      @package_options = options
      @package = Axlsx::Package.new
      instance_exec &block
    end

    def spreadsheet options = {}, &block
      if block_given?
        @package.workbook.add_worksheet do |spreadsheet|
          @spreadsheet = spreadsheet
          @spreadsheet.name = options[:name] unless options[:name].respond_to?(:empty?) ? options[:name].empty? : !options[:name]
          build_spreadsheet_styles
          instance_exec &block
        end
      else
        build_spreadsheet
      end
    end
    alias_method :table, :spreadsheet

    def build_spreadsheet
      @spreadsheet ||= @package.workbook.add_worksheet.tap { @row = @cell = nil }
    end

    def row options = {}, &block
      if block_given?
        @row_options = options
        spreadsheet.add_row do |row|
          @row = row
          instance_exec &block
        end
      else
        build_row
      end
    end
    alias_method :tr, :row

    def build_row
      @row ||= spreadsheet.add_row.tap{ @cell = nil }
    end

    def cell value, options = {}
      unless @row_options.respond_to?(:empty?) ? @row_options.empty? : !@row_options
        options = @row_options.merge(options){ |key, row_option, cell_option| Array.new << row_option << cell_option }
      end
      @cell = row.add_cell value.to_s, style: style_from_options(options)
      unless options[:colspan].respond_to?(:empty?) ? options[:colspan].empty? : !options[:colspan]
        spreadsheet.merge_cells "#{pos[:x]}#{pos[:y]}:#{shift_x(pos[:x], 5)}#{pos[:y]}"
      end
    end
    alias_method :td, :cell

    def save file_name
      @package.serialize file_name
      File.open file_name
    end

    def to_temp_file
      stream = @package.to_stream
      Tempfile.new(%w(temporary-workbook .xlsx), encoding: 'utf-8').tap do |file|
        file.write stream.read
        file.close
      end
    end

    def to_xml
      spreadsheet.to_xml_string
    end

    private

      def pos
        { x: x_axis[@cell.pos[0]], y: y_axis[@cell.pos[1]] }
      end

      def shift_x x, diff
        x_axis[x_axis.index(x) + diff]
      end

      def style_from_options options
        if options[:style].is_a? Array
          add_complex_style options[:style]
        end
        spreadsheet_style_by_name style_name(options[:style])
      end

      def style_name style_keys
        Array(style_keys).join('_').to_sym
      end

      def add_complex_style style_keys
        complex_option_args = {}
        style_keys.map{ |style_key| style_builder_args[style_key] }.compact.map{ |args| complex_option_args.merge! args }
        @style_builder_args[style_name(style_keys)] = complex_option_args
        rebuild_spreadsheet_styles
      end

      def style_builder_args
        @style_builder_args ||= {
            center: { alignment: { horizontal: :center } },
            bold:   { b: true },
            italic: { i: true }
        }
      end

      def build_spreadsheet_styles
        @spreadsheet_styles = {}
        spreadsheet.workbook.styles do |style_builder|
          style_builder_args.each do |style_name, style_args|
            @spreadsheet_styles[style_name] = style_builder.add_style style_args
          end
        end
        @spreadsheet_styles
      end
      alias_method :rebuild_spreadsheet_styles, :build_spreadsheet_styles

      def spreadsheet_styles
        @spreadsheet_styles ||= build_spreadsheet_styles
      end

      def spreadsheet_style_by_name name
        spreadsheet_styles[name]
      end

      # should be dynamic but for now to avoid calculations is static
      def x_axis
        ('A'..'Z').to_a
      end

      # should be dynamic but for now to avoid calculations is static
      def y_axis
        (1..1000).to_a
      end
  end
end