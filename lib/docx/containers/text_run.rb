require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class TextRun
        include Container
        include Elements::Element

        DEFAULT_FORMATTING = {
          italic:    false,
          bold:      false,
          underline: false,
          strike: false
        }

        def self.tag
          'r'
        end

        attr_reader :text
        attr_reader :formatting

        def initialize(node, document_properties = {})
          @node = node
          @text_nodes = @node.xpath('w:t|w:r/w:t').map { |t_node| Elements::Text.new(t_node) }
          @drawings = @node.xpath('w:drawing').map { |t_node| Containers::Drawing.new(t_node, document_properties) }
          @objects = @node.xpath('w:object').map { |t_node| Containers::BinObject.new(t_node, document_properties) }

          @properties_tag = 'rPr'
          @text       = parse_text || ''
          @formatting = parse_formatting || DEFAULT_FORMATTING
          @document_properties = document_properties
          @font_size = @document_properties[:font_size]
        end

        # Set text of text run
        def text=(content)
          if @text_nodes.size == 1
            @text_nodes.first.content = content
          elsif @text_nodes.empty?
            new_t = Elements::Text.create_within(self)
            new_t.content = content
          end
          reset_text
        end

        # Returns text contained within text run
        def parse_text
          @text_nodes.map(&:content).join('')
        end

        # Substitute text in text @text_nodes
        def substitute(match, replacement)
          @text_nodes.each do |text_node|
            text_node.content = text_node.content.gsub(match, replacement)
          end
          reset_text
        end

        def parse_formatting
          {
            italic: @node.at_xpath('.//w:i') && @node.at_xpath('.//w:i').attributes['val']&.value != '0',
            bold: @node.at_xpath('.//w:b') && @node.at_xpath('.//w:b').attributes['val']&.value != '0',
            underline: !@node.xpath('.//w:u').empty?,
            strike: !@node.xpath('.//w:strike').empty?
          }
        end

        def to_s
          @text
        end

        # Return text as a HTML fragment with formatting based on properties.
        def to_html
          html = @text
          html = HTML.content_tag(:em, html) if italicized?
          html = HTML.content_tag(:strong, html) if bolded?
          html = HTML.content_tag(:s, html) if striked?
          styles = []
          styles << 'text-decoration: underline' if underlined?
          # No need to be granular with font size down to the span level if it doesn't vary.
          # styles << "font-size :#{font_size}pt" if font_size != @font_size
          html = HTML.content_tag(:span, html, styles: styles.join(';')) unless styles.empty?
          html = HTML.content_tag(:a, html, href:, target: '_blank') if hyperlink?
          html = HTML.join([html, *HTML.join(@drawings.map(&:to_html))]) if @drawings
          html = HTML.join([html, *HTML.join(@objects.map(&:to_html))]) if @objects
          return html
        end

        def italicized?
          @formatting[:italic]
        end

        def bolded?
          @formatting[:bold]
        end

        def striked?
          @formatting[:strike]
        end

        def underlined?
          @formatting[:underline]
        end

        def hyperlink?
          @node.name == 'hyperlink' && external_link?
        end

        def external_link?
          !@node.attributes['id'].nil?
        end

        def href
          @document_properties[:hyperlinks][hyperlink_id]
        end

        def hyperlink_id
          @node.attributes['id'].value
        end

        def font_size
          size_attribute = @node.at_xpath('w:rPr//w:sz//@w:val')

          return @font_size unless size_attribute

          size_attribute.value.to_i / 2
        end

        private

        def reset_text
          @text = parse_text
        end
      end
    end
  end
end
