require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class Custom
      include Powerpoint::Util

      attr_reader :title, :subtitle, :properties, :coords

      def initialize(options={})
        require_arguments [:title, :properties], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @coords = default_coords unless @coords.present?
        properties
      end

      def save(extract_path, index)
        properties.each do |property|
          if property[:type] == 'image'
            copy_media(extract_path, property[:file_path]) 
          end
        end
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def file_type
        File.extname(image_name).gsub('.', '')
      end

      def default_coords
        {x: 0, y: 0, cx: pixle_to_pt(50), cy: pixle_to_pt(50)}
      end
      private :default_coords

      def save_rel_xml(extract_path, index)
        render_view('custom_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels", index: index)
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('custom.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end