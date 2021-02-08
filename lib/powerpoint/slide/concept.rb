require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class Concept
      include Powerpoint::Util

      attr_reader :title, :subtitle,:subtitle_2, :logo, :image_information, :images

      def initialize(options={})
        require_arguments [:title, :subtitle,:subtitle_2, :logo, :image_information, :images], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @images = images
         
      end

      def save(extract_path, index)
        @images.each do |image|
          copy_media(extract_path, image)
        end
        
        copy_media(extract_path, logo)

        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def image_ratio(image_information)
        if image_information[0] > image_information[1]
          image_ratio = ['width',image_information[2]]
        else
          image_ratio = ['height',image_information[2]]
        end
      end

      def save_rel_xml(extract_path, index)
        render_view('concept_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels", index: index)
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('concept.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end