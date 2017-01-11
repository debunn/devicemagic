require 'google/apis/sheets_v4'
require 'googleauth'
require 'googleauth/stores/file_token_store'

require 'fileutils'
require 'json'

OOB_URI = 'urn:ietf:wg:oauth:2.0:oob'
APPLICATION_NAME = 'SheetsTest'
CLIENT_SECRETS_PATH = 'client_secret.json'
CREDENTIALS_PATH = File.join(Dir.home, '.credentials',
                             "sheets.googleapis.com-ruby-quickstart.yaml")
SCOPE = Google::Apis::SheetsV4::AUTH_SPREADSHEETS_READONLY

##
# Ensure valid credentials, either by restoring from the saved credentials
# files or intitiating an OAuth2 authorization. If authorization is required,
# the user's default browser will be launched to approve the request.
#
# @return [Google::Auth::UserRefreshCredentials] OAuth2 credentials
def authorize
  FileUtils.mkdir_p(File.dirname(CREDENTIALS_PATH))

  client_id = Google::Auth::ClientId.from_file(CLIENT_SECRETS_PATH)
  token_store = Google::Auth::Stores::FileTokenStore.new(file: CREDENTIALS_PATH)
  authorizer = Google::Auth::UserAuthorizer.new(
      client_id, SCOPE, token_store)
  user_id = 'default'
  credentials = authorizer.get_credentials(user_id)
  if credentials.nil?
    url = authorizer.get_authorization_url(
        base_url: OOB_URI)
    puts "Open the following URL in the browser and enter the " +
             "resulting code after authorization"
    puts url
    code = gets
    credentials = authorizer.get_and_store_credentials_from_code(
        user_id: user_id, code: code, base_url: OOB_URI)
  end
  credentials
end

# Initialize the API
service = Google::Apis::SheetsV4::SheetsService.new
service.client_options.application_name = APPLICATION_NAME
service.authorization = authorize

# This is the Google Sheets ID for my test spreadsheet
spreadsheet_id = '1-Nt8T7xVQpX0xPQHzTZmEnHU1hRB_0iZ9LEjalIzXN0'

# Get the spreadsheet properties to determine each sheet name
response = service.get_spreadsheet(spreadsheet_id)

masterhash = {} # The top level hash which holds all spreadsheet data

# Run through each sheet individually
response.sheets.each do |sheet|
  # Query this sheet for all data, then parse in to a hash
  sheetdata = service.get_spreadsheet_values(spreadsheet_id, sheet.properties.title)

  titledata = [] # the array which holds the column header titles
  sheetarr = [] # the array which holds each row's data

  # Run through each row in the sheet, adding values to the valuehash
  sheetdata.values.each_with_index do |row, index|
    if index == 0
      # Set the title values for this sheet
      titledata = row
    else
      # Loop through each subsequent row, and add the values to the valuehash
      valuehash = {}
      row.each_with_index do |cell, index|
        valuehash[titledata[index]] = cell
      end

      # Add this row of values to this sheets value array
      sheetarr.push valuehash
    end
  end

  # Add this sheet to the top level hash
  masterhash[sheet.properties.title] = sheetarr
end

# Convert the hash to JSON, output the JSON to stdout
puts masterhash.to_json
