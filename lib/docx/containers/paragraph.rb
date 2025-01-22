require 'docx/containers/text_run'
require 'docx/containers/container'
require 'html'

module Docx
  module Elements
    module Containers
      class Paragraph
        include Container
        include Elements::Element

        TEXT_RUN_NODES = 'm:oMath|m:oMathPara|w:r|w:hyperlink'.freeze

        def self.tag
          'p'
        end

        # Child elements: pPr, r, fldSimple, hlink, subDoc
        # http://msdn.microsoft.com/en-us/library/office/ee364458(v=office.11).aspx
        def initialize(node, document_properties = {}, doc = nil)
          @node = node
          @properties_tag = 'pPr'
          @document_properties = document_properties
          @font_size = @document_properties[:font_size]
          @document = doc
        end

        # Set text of paragraph
        def text=(content)
          if text_runs.size == 1
            text_runs.first.text = content
          elsif text_runs.size == 0
            new_r = TextRun.create_within(self)
            new_r.text = content
          else
            text_runs.each {|r| r.node.remove }
            new_r = TextRun.create_within(self)
            new_r.text = content
          end
        end

        # Return text of paragraph
        def to_s
          text_runs.map(&:text).join('')
        end

        # Return paragraph as a <p></p> HTML fragment with formatting based on properties.
        def to_html
          html = HTML.join(text_runs.map(&:to_html))
          styles = {}
          styles['name'] = style if style_id && !style&.empty?
          styles['font-size'] = "#{font_size}pt" if size_attribute
          styles['color'] = "##{font_color}" if font_color
          styles['text-align'] = alignment if alignment

          return HTML.content_tag(:ul, HTML.content_tag(:li, html, styles)) if list_item_level

          HTML.content_tag(:p, html, styles)
        end

        # Array of text runs contained within paragraph
        def text_runs
          @node.xpath(TEXT_RUN_NODES).map do |r_node|
            case r_node.name
            when 'r', 'hyperlink'
              next Containers::TextRun.new(r_node, @document_properties)
            when 'oMathPara'
              next Containers::Math.new(r_node, false, @document_properties)
            when 'oMath'
              next Containers::Math.new(r_node, true, @document_properties)
            end
          end
        end

        # Iterate over each text run within a paragraph
        def each_text_run
          text_runs.each { |tr| yield(tr) }
        end

        def aligned_left?
          ['left', nil].include?(alignment)
        end

        def aligned_right?
          alignment == 'right'
        end

        def aligned_center?
          alignment == 'center'
        end

        def size_attribute
          @node.at_xpath('w:pPr//w:sz//@w:val')
        end

        def font_size
          return @font_size unless size_attribute

          size_attribute.value.to_i / 2
        end

        def font_color
          color_tag = @node.xpath('w:r//w:rPr//w:color').first || @node.xpath('w:pPr//w:rPr//w:color').first
          color_tag ? color_tag.attributes['val'].value : nil
        end

        def list_item_level
          @node.at_xpath('w:pPr//w:numPr//w:ilvl//@w:val')
        end

        def list_id
          @node.at_xpath('w:pPr//w:numPr//w:numId//@w:val')&.value
        end

        def style
          return nil unless @document

          @document.style_name_of(style_id) ||
            @document.default_paragraph_style
        end

        def style_id
          style_property&.get_attribute('w:val')
        end

        def style=(identifier)
          id = @document.styles_configuration.style_of(identifier).id

          style_property.set_attribute('w:val', id)
        end

        alias_method :style_id=, :style=
        alias_method :text, :to_s

        private

        def style_property
          properties&.at_xpath('w:pStyle') || properties&.add_child('<w:pStyle/>')&.first
        end

        # Returns the alignment if any, or nil if left
        def alignment
          @node.at_xpath('.//w:jc/@w:val')&.value
        end
      end
    end
  end
end
