module ActiveAdmin
  module Axlsx
    module ResourceControllerExtension
      def self.included(base)
        # base.send :alias_method, :per_page_without_xlsx, :xlsx
        # base.send :alias_method, :xlsx, :per_page_with_xlsx
        base.send :respond_to, :xlsx, only: :index
        base.send :respond_to, :xlsx
      end

      # patching the index method to allow the xlsx format.
      # Patches index to respond to requests with xls mime type by
      # sending a generated xls document serializing the current
      # collection
      def index
        super do |format|
          format.xlsx do
            xlsx = active_admin_config.xlsx_builder.serialize(collection)
            send_data(xlsx,
                      :filename  => "#{xlsx_filename}",
                      type: Mime::Type.lookup_by_extension(:xlsx))
          end

          yield(format) if block_given?
        end
      end

      # patching per_page to use the CSV record max for pagination when the format is xlsx
      # def per_page_with_xlsx
      #   if request.format ==  Mime::Type.lookup_by_extension(:xlsx)
      #     per_page do
      #       return max_csv_records
      #     end
      #   end
      # end

      # Returns a filename for the xlsx file using the collection_name
      # and current date such as 'my-articles-2011-06-24.xlsx'.
      def xlsx_filename
        "#{resource_collection_name.to_s.gsub('_', '-')}-#{Time.now.strftime("%Y-%m-%d")}.xlsx"
      end
    end
  end
end
