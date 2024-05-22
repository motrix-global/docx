require 'docx/containers/table_cell'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class TableRow
        include Container
        include Elements::Element

        def self.tag
          'tr'
        end

        def initialize(node, document_properties = {})
          @node = node
          @properties_tag = ''
          @document_properties = document_properties
        end

        # Array of cells contained within row
        def cells
          @node.xpath('w:tc').map { |c_node| Containers::TableCell.new(c_node, @document_properties) }
        end

      end
    end
  end
end
