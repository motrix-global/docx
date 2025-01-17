require 'docx/containers/container'
require 'mathtype_to_mathml'
require 'mathml2latex'
require 'tempfile'

module Docx
  module Elements
    module Containers
      class BinObject
        include Container
        include Elements::Element

        def self.tag
          'object'
        end

        def initialize(node, document_properties = {})
          @node = node
          @document_properties = document_properties
        end

        def to_html
          begin
            ole_binary_file = @document_properties[:doc_bundle].glob("word/#{src}").first
            file = Tempfile.new(src.split(/(?=\.)/))
            file.write(ole_binary_file.get_input_stream.read)
            file.rewind

            mathml = MathTypeToMathML::Converter.new(file.path).convert

            latex = Mathml2latex::Converter.new.to_latex(mathml)
          rescue
            STDERR.puts "Erro processando #{src}"
          ensure
            file&.close
            file&.unlink
          end

          HTML.content_tag('inline-math', latex.to_s, type: 'latex')
        end

        def src
          @document_properties[:objects][bin_object_id]
        end

        def bin_object_id
          @node.at_xpath('.//*:OLEObject').attributes['id'].value
        end
      end
    end
  end
end
