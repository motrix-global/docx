require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class Drawing
        include Container
        include Elements::Element

        def self.tag
          'drawing'
        end

        def initialize(node, document_properties = {})
          @node = node
          @document_properties = document_properties
        end

        def to_html
          html_tag(:img, attributes: { src: src })
        end

        def src
          @document_properties[:images][image_id]
        end

        def image_id
          @node.at_xpath('.//*:blip').attributes['embed'].value
        end
      end
    end
  end
end
