require 'docx/containers/text_run'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class TableCell
        include Container
        include Elements::Element

        def self.tag
          'tc'
        end

        def initialize(node, document_properties = {})
          @node = node
          @properties_tag = 'tcPr'
          @document_properties = document_properties
        end

        # Return text of paragraph's cell
        def to_s
          paragraphs.map(&:text).join('')
        end

        # Array of paragraphs contained within cell
        def paragraphs
          @node.xpath('w:p').map { |p_node| Containers::Paragraph.new(p_node, @document_properties) }
        end

        def split_p
          word_paragraphs = @node.xpath('w:p').map do |paragraph|
            splitted_runs = paragraph.xpath(Paragraph::TEXT_RUN_NODES).slice_when { |prev, _| prev.at_xpath('w:br') }.to_a
            paragraph.xpath(Paragraph::TEXT_RUN_NODES).each(&:remove)

            splitted_runs.map do |runs_group|
              new_paragraph = paragraph.dup
              runs_group.each { |r| new_paragraph.add_child(r) }
              new_paragraph
            end
          end.flatten

          word_paragraphs.map { |p_node| Containers::Paragraph.new(p_node, @document_properties) }
        end

        # Iterate over each text run within a paragraph's cell
        def each_paragraph
          paragraphs.each { |tr| yield(tr) }
        end

        def to_html(header:)
          cell_tag = header ? :th : :td
          HTML.content_tag(cell_tag, HTML.join(paragraphs.map { |p| HTML.join(p.text_runs.map(&:to_html)) }))
        end

        alias_method :text, :to_s
      end
    end
  end
end
