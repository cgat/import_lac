def unbundled_require(gem_name)
  if defined?(::Bundler)
    spec_path = Dir.glob("/Users/mlp/.rvm/gems/ruby-1.9.3-p194/specifications/#{gem_name}-*.gemspec").last
    if spec_path.nil?
      warn "Couldn't find #{gem_name}"
      return
    end

    spec = Gem::Specification.load spec_path
    spec.activate
  end

  begin
    require gem_name
    yield if block_given?
  rescue Exception => err
    warn "Couldn't load #{gem_name}: #{err}"
  end
end

unbundled_require 'ruby-ole'
unbundled_require 'spreadsheet'

module Import
  def import_images(spreadsheet_path, spreadsheet_output_path, sheet_names, postives_dir, envelopes_dir)
    sheet_names = Array(sheet_names)
    files_list = Dir.glob(File.join(postives_dir, "**","*"))
    envelope_files_list = Dir.glob(File.join(envelopes_dir, "**","*"))
    book = Spreadsheet.open(spreadsheet_path)

    import_objects = load_import_objects(book,sheet_names, files_list, envelope_files_list)
    import_objects.each{|imp_obj| imp_obj.import unless imp_obj.is_envelope}
    book.write(spreadsheet_output_path)
    nil
  end

  def load_import_objects(spreadsheet, sheet_names, files_list, envelope_files_list)
    import_objects = [];
    spreadsheet.worksheets.each do |wsheet|
      if sheet_names.include?(wsheet.name)
        wsheet.each 2 do |row|
          unless row[0].blank? && row[1].blank? && row[2].blank?
            begin
              imp_obj =  ImportObject.new(row)
              imp_obj.image_path = files_list.find{|f| f=~/#{imp_obj.camera_id}.jpg$/}
              if imp_obj.image_path.blank?
                imp_obj.image_path = envelope_files_list.find{|f| f=~/#{imp_obj.camera_id}.jpg$/}
                if img_obj.image_path.blank?
                  raise StandardError, "No image_path found in list #{img_obj.row}"
                else
                  img_obj.is_envelope=true
                end
              end
              import_objects<< imp_obj
            rescue => e
              puts "Error: #{row} : #{e}"
            end
          end
        end
      end
    end
    import_objects
  end




  class ImportObject
    attr_accessor :row, :surveyor, :survey, :survey_season, :station_name, :plate_id, :camera_id, :comments, :box, :image_path, :is_envelope
    def initialize(row)
      self.row = row
      self.surveyor = Surveyor.find_by_last_name!(row[0])
      self.survey = surveyor.surveys.find_by_name!(row[1])
      self.survey_season = survey.survey_seasons.find_by_year!([2].chomp(".0"))
      self.station_name = row[3].chomp(".0")
      self.plate_id = row[4].chomp(".0")
      self.camera_id = convert_camera_id(row[5])
      self.comments = row[6]
      self.box = row[9]
      self.is_envelope = false
    end

    def convert_camera_id(cam_id)
      id = cam_id[/\d{3}-{0,1}\d{4}/]
      raise StandardError, "Camera_id did not match pattern" if id.blank?
      id = "P#{id.gsub("-","")}"
    end

    def import
      raise StandardError, "You must have a image file associated with the import object before importing #{row}" if image_path.blank?
      raise StandardError, "Can't import envelope files" if is_envelope
      if station_name.present?
        station = survey_season.stations.where(name: station_name).first_or_create
        historic_visit = station.historic_visit || station.historic_visit.create()
        parent = historic_visit
      else
        parent = survey_season
      end
      hcap = HistoricCapture.where(field_note_photo_reference: plate_id, plate_id: plate_id, comments: comments, capture_owner_id: parent.id, capture_owner_type: parent.class.name).first_or_create
      if hcap.capture_images.where("image like '%#{camera_id}'").blank?
        hcap.capture_images.create(image: File.open(image_path), image_state: "MISC")
      end
      row[8] = "http://envi-mountain-0003.envi.uvic.ca/historic_captures/#{hcap.id}"
    end
  end

end
