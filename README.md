[![version](https://badge.fury.io/rb/powerpoint.svg)](https://badge.fury.io/rb/powerpointk)
[![downloads](https://ruby-gem-downloads-badge.herokuapp.com/powerpoint?type=total&total_label=downloads)](https://ruby-gem-downloads-badge.herokuapp.com/powerpoint?type=total&total_label=downloads)

# 'powerpoint-pro' gem -- for creating PowerPoint Slides in Ruby.

'powerpoint-pro' is a Ruby gem that can generate PowerPoint files(pptx).

## Installation

'powerpoint-pro' can be used from the command line or as part of a Ruby web framework. To install the gem using terminal, run the following command:

    gem install "powerpoint-pro"

To use it in Rails, add this line to your Gemfile:

    gem "powerpoint-pro"


## Basic Usage

'powerpoint-pro' gem can generate a PowerPoint presentaion based on a standard template:

```ruby
require 'powerpoint'

@deck = Powerpoint::Presentation.new

# Creating an introduction slide:
title = 'Bicycle Of the Mind'
subtitle = 'created by Steve Jobs'
@deck.add_intro title, subtitle

# Creating a text-only slide:
# Title must be a string.
# Content must be an array of strings that will be displayed as bullet items.
title = 'Why Mac?'
content = ['Its cool!', 'Its light.']
@deck.add_textual_slide title, content

# Creating an image Slide:
# It will contain a title as string.
# and an embeded image
title = 'Everyone loves Macs:'
image_path = 'samples/images/sample_gif.gif'
@deck.add_pictorial_slide title, image_path

# Specifying coordinates and image size for an embeded image.
# x and y values define the position of the image on the slide.
# cx and cy define the width and height of the image.
# x, y, cx, cy are in points. Each pixel is 12700 points.
# coordinates parameter is optional.
coords = {x: 124200, y: 3356451, cx: 2895600, cy: 1013460}
@deck.add_pictorial_slide title, image_path, coords

# Saving the pptx file to the current directory.
@deck.save('test.pptx')
```

## Custom Slide
```ruby
@deck = Powerpoint::Presentation.new

# Creating an Custom slide:
title = 'Bicycle Of the Mind'
subtitle = 'created by Steve Jobs'
components = [
    {type: "image", title: "image 2", file_path: '/path', coords: {x: 750, y: 0, cx: 50, cy: 50}}, 
    {type: "text", content: "It Is My Text", size: 14, bold: true, align: 'right', font: "Snell Roundhand"},
    ]
# Creating a Custom slide text, image every thing can be configurable:
@deck.add_add_custom_slide title, subtitle, components
@deck.save('test.pptx')
```
  
## Compatibility

'powerpoint-pro' gem has been tested with LibreOffice (4.2.1.1) and Apache OpenOffice (4.0.1) on Mac OS X Mavericks, Microsoft PowerPoint 2010 on Windows 7 and Google Docs (latest version as of Jan 2023).

## Contributing

Contributions are welcomed. You can fork a repository, add your code changes to the forked branch, ensure all existing unit tests pass, create new unit tests cover your new changes and finally create a pull request.

After forking and then cloning the repository locally, install Bundler and then use it to install the development gem dependecies:

    gem install bundler
    bundle install

Once this is complete, you should be able to run the test suite:

    rake


## Bug Reporting

Please use the [Issues](https://github.com/prosenjit98/powerpoint_pro/issues) page to report bugs or suggest new enhancements.

## License

Powerpoint has been published under [MIT License](https://github.com/prosenjit98/powerpoint_pro/blob/master/LICENSE.txt)
