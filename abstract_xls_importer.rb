#encoding: utf-8

class AbstractXlsImporter
  class ImportError < StandardError; end
  class InvalidCollection < ImportError
    def description
      'Collection invalide'
    end
  end
  class InvalidCollaboration < ImportError
    def description
      'Collaboration invalide'
    end
  end
  class InvalidSubCategory < ImportError
    def description
      'Catégorie invalide'
    end
  end
  class NoHeader < ImportError
    def description; 'La première ligne du fichier doit comporter le nom des colonnes'; end
  end
  class InvalidHeader < ImportError
    def description
      'Nom de colonne inconnu'
    end
  end

  extend ActiveModel::Naming
  include ActiveModel::Conversion
  include ActiveModel::AttributeMethods
  include ActiveModel::Validations

  attr_accessor :country_id, :xls
  attr_reader :invalid_rows, :imported_count, :total_rows

  validates_presence_of :xls

  def initialize(attributes = {})
    @xls = if attributes[:xls].respond_to?(:tempfile)
      attributes[:xls].tempfile
    else
      attributes[:xls]
    end
    @invalid_rows = {}
    @imported_count = 0
  end

  def header_column_names
    [ ]
  end

  def header_main_column_name
    ""
  end

  def import
    begin
      xls_hashes = create_xls
    rescue ImportError => e
      @invalid_rows[e] ||= []
    end
    return unless @invalid_rows.empty?

    @total_rows = xls.size
    @total_rows = xls_hashes.size

    xls_hashes.each do |row|
      begin
        import_from_xls_row row
      rescue ImportError => e
        @invalid_rows[e] ||= []
        @invalid_rows[e] << row
      end
    end
  end

  def errors_count
    @invalid_rows.values.flatten.size
  end

  def persisted?
    false
  end

  def create_xls
    b = Spreadsheet.open @xls
    s = b.worksheets.first

    column_names = s.row(0).to_a.map { |header| header_to_column_name header }
    fail NoHeader if column_names.empty?

    hashes = []
    s.each 1 do |row|
      hashes << xls_row_to_hash(column_names, row)
    end

    hashes.select! { |hash| hash[header_main_column_name].present? }
    hashes
 end

  def xls_row_to_hash(column_names, row)
    h = {}
    row_array = row.to_a

    column_names.each_with_index do |column_name, index|
      if column_name
        h[column_name] = row_array[index]
        h[column_name] = h[column_name].to_i.to_s if row_array[index].class == Float
      end
    end
    h
  end

  def import_from_xls_row( row )
  end

  include ActionView::Helpers
  def to_html(description)
    return "" if description.blank?
    description.gsub(/\\n/, "\n")
  end

  def header_to_column_name(header)
    return "" unless header
     # underscore function does not convert spaces, using gsub
     # header = header.downcase.gsub(/\s/,"_")
    header if header_column_names.include? header
   end
 end


