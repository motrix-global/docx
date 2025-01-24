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

          @node.xpath('//m:oMathParaPr').each(&:remove)

          if @node.at_xpath('m:acc')
            acc_type = @node.at_xpath('m:acc/m:accPr/m:chr')
            if acc_type&.attribute('val')&.value == '̅'
              @node.xpath('m:acc').each { |n| n.name = 'bar' }
            elsif !acc_type
              @node.xpath('m:acc').each do |n|
                n.name = 'limUpp'
                lim_node = Nokogiri::XML::Node.new('m:lim', @node.document)
                r_node = Nokogiri::XML::Node.new('m:r', @node.document)
                t_node = Nokogiri::XML::Node.new('m:t', @node.document)
                t_node.inner_html = '&#x302;' # Unicode for hat accent (^)

                # Assemble the structure
                r_node.add_child(t_node)    # Add <m:t> inside <m:r>
                lim_node.add_child(r_node)  # Add <m:r> inside <m:lim>
                n.add_child(lim_node)
              end
              # @node.xpath('//m:accPr').each(&:remove)
            end
          end

          @formula = Plurimath::Math.parse(@node.to_xml, :omml)
        rescue StandardError => e
          puts e
        end

        def to_html
          HTML.content_tag(math_tag, text, type: 'latex')
        end

        def math_tag
          @inline ? 'inline-math' : 'math-block'
        end

        def text
          return HTML::Fragment.new('ERROR') unless @formula

          latex_formula = @formula.to_latex.gsub('<', ' \lt ').gsub('>', ' \gt ').gsub(/[[:space:]]+/, ' ').strip
          HTML::Fragment.new(latex_formula)
        rescue StandardError => e
          puts e
          'ERROR'
        end
      end
    end
  end
end
