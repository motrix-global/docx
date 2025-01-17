require 'docx/containers/table_row'
require 'docx/containers/table_column'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class Table
        include Container
        include Elements::Element

        def self.tag
          'tbl'
        end

        def initialize(node, document_properties = {})
          @node = node
          @properties_tag = 'tblGrid'
          @document_properties = document_properties
        end

        def to_html
          table_header = HTML.content_tag(:thead, rows[0].to_html(header: true))
          table_body = HTML.content_tag(:tbody, HTML.join(rows[1..].map(&:to_html)))
          HTML.content_tag(:table, HTML.join([table_header, table_body]))
        end

        # Array of row
        def rows
          @node.xpath('w:tr').map { |r_node| Containers::TableRow.new(r_node, @document_properties) }
        end

        def row_count
          @node.xpath('w:tr').count
        end

        # Array of column
        def columns
          columns_containers = []
          (0..(column_count-1)).each do |i|
            columns_containers[i] = Containers::TableColumn.new @node.xpath("w:tr//w:tc[#{i+1}]")
          end
          columns_containers
        end

        def column_count
          @node.xpath('w:tblGrid/w:gridCol').count
        end

        # Iterate over each row within a table
        def each_rows
          rows.each { |r| yield(r) }
        end

      end
    end
  end
end
