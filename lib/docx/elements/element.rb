require 'nokogiri'
require 'docx/elements'
require 'docx/containers'
require 'html'

module Docx
  module Elements
    module Element
      extend HTML
      DEFAULT_TAG = ''

      # Ensure that a 'tag' corresponding to the XML element that defines the element is defined
      def self.included(base)
        base.extend(ClassMethods)
        base.const_set(:TAG, Element::DEFAULT_TAG) unless base.const_defined?(:TAG)
      end

      attr_accessor :node

      # TODO: Should create a docx object from this
      def parent(type = '*')
        @node.at_xpath("./parent::#{type}")
      end

      def at_xpath(*args)
        @node.at_xpath(*args)
      end

      def xpath(*args)
        @node.xpath(*args)
      end

      # Get parent paragraph of element
      def parent_paragraph
        Elements::Containers::Paragraph.new(parent('w:p'))
      end

      # Insertion methods
      # Insert node as last child
      def append_to(element)
        @node = element.node.add_child(@node)
        self
      end

      # Insert node as first child (after properties)
      def prepend_to(element)
        @node = element.node.properties.add_next_sibling(@node)
        self
      end

      def insert_after(element)
        # Returns newly re-parented node
        @node = element.node.add_next_sibling(@node)
        self
      end

      def insert_before(element)
        @node = element.node.add_previous_sibling(@node)
        self
      end

      # Creation/edit methods
      def copy
        self.class.new(@node.dup)
      end

      module ClassMethods
        def create_with(element)
          # Need to somehow get the xml document accessible here by default, but this is alright in the interim
          self.new(Nokogiri::XML::Node.new("w:#{self.tag}", element.node.document))
        end

        def create_within(element)
          new_element = create_with(element)
          new_element.append_to(element)
          new_element
        end
      end
    end
  end
end
