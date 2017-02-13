require 'roo/excelx/extractor'

module Roo
  class Excelx
    class Workbook < Excelx::Extractor
      class Label
        attr_reader :sheet, :row, :col, :name

        def initialize(name, sheet, row, col)
          @name = name
          @sheet = sheet
          @row = row.to_i
          @col = ::Roo::Utils.letter_to_number(col)
        end

        def key
          [@row, @col]
        end
      end

      def initialize(path)
        super
        fail ArgumentError, 'missing required workbook file' unless doc_exists?
      end

      def sheets
        doc.xpath('//sheet')
      end

      # aka labels
      def defined_names
        Hash[doc.xpath('//definedName').map do |defined_name|
          name = defined_name['name']

          # non-adjacent named cells are separated via comma
          [name] << defined_name.text.split(',').reduce([]) { |memo, range|
            # each 'range', where a range could just be a single cell, is separated into
            # two parts by '!$'.
            memo << range.split('!$').each_slice(2).collect { |(sheet, coordinates)|
              # a range will use ':$' to separate the start cell from the end cell. a
              # single cell will not have the colon.
              [sheet] << coordinates.split(':$').map { |cell|
                # each cell has a column and row identifier, separated by a '$'
                cell.split('$')
              }
            }.reduce([]) { |memo, (sheet, ((x1, x2), (y1, y2)))|
              # a named range with a single cell will not have a second range
              y1 ||= x1
              y2 ||= x2

              # create an array containing a range for columns and rows
              memo << [sheet, (x1..y1), (x2..y2)]
            }.reduce([]) { |memo, (sheet, cols, rows)|
              # iterate over each column/row combination to create a label
              memo << cols.to_a.product(rows.to_a).map { |col, row|
                Label.new(name, sheet, row, col)
              }
            }
          # flatten nested arrays, occurs with non-adjacent named cells
          }.flatten
        end]
      end

      def base_date
        @base_date ||=
        begin
          # Default to 1900 (minus one day due to excel quirk) but use 1904 if
          # it's set in the Workbook's workbookPr
          # http://msdn.microsoft.com/en-us/library/ff530155(v=office.12).aspx
          result = Date.new(1899, 12, 30) # default
          doc.css('workbookPr[date1904]').each do |workbookPr|
            if workbookPr['date1904'] =~ /true|1/i
              result = Date.new(1904, 01, 01)
              break
            end
          end
          result
        end
      end
    end
  end
end
