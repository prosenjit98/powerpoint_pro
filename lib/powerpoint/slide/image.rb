require 'zip/filesystem'
require 'fileutils'
require 'fastimage'
require 'erb'

module Powerpoint
  module Slide
    class Image
      include Powerpoint::Util

      attr_reader :title, :subtitle, :images, :coords

      def initialize(options={})
        require_arguments [:title, :subtitle, :images], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @coords = default_coords unless @coords.present?
        images
      end

      def save(extract_path, index)
        images.each do |image|
          if image[:type] == 'image'
            copy_media(extract_path, image[:file_path])
          end
        end
        save_rel_xml(extract_path, index)
        save_slide_xml(extract_path, index)
      end

      def file_type
        File.extname(image_name).gsub('.', '')
      end

      def default_coords
        # slide_width = pixle_to_pt(720)
        # default_width = pixle_to_pt(550)

        # return {} unless dimensions = FastImage.size(image_path)
        # image_width, image_height = dimensions.map {|d| pixle_to_pt(d)}
        # new_width = default_width < image_width ? default_width : image_width
        # ratio = new_width / image_width.to_f
        # new_height = (image_height.to_f * ratio).round
        {x: 0, y: 0, cx: pixle_to_pt(50), cy: pixle_to_pt(50)}
      end
      private :default_coords

      def save_rel_xml(extract_path, index)
        render_view('image_rel.xml.erb', "#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels", index: index)
      end
      private :save_rel_xml

      def save_slide_xml(extract_path, index)
        render_view('image.xml.erb', "#{extract_path}/ppt/slides/slide#{index}.xml")
      end
      private :save_slide_xml
    end
  end
end