require 'docx/containers/container'
require 'mathtype_to_mathml'
require 'mathml2latex'
require 'plurimath'
require 'tempfile'

module Docx
  module Elements
    module Containers
      class Math
        include Container
        include Elements::Element

        def self.tag
          'math'
        end

        attr_reader :inline

        def initialize(node, inline, document_properties = {})
          @node = node # (inline ? node : node.xpath('m:oMath').first)
          @document_properties = document_properties
          @inline = inline

          @formula = Plurimath::Math.parse(@node.to_xml, :omml)
        end

        def to_html
          HTML.content_tag(math_tag, text, type: 'latex')
        end

        def math_tag
          @inline ? 'inline-math' : 'math-block'
        end

        def text
          latex_formula = @formula.to_latex.gsub('<', '\lt').gsub('>', '\gt').gsub(/[[:space:]]/, ' ')
          HTML::Fragment.new(latex_formula)
        rescue StandardError => e
          puts e
          'ERROR'
        end
      end
    end
  end
end
