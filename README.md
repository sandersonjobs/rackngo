# RackNGo Spreadsheet Processor

RackNGo is a spreadsheet processing tool used to enter devices and equipment into the Metal API

## Set Up

### Set Environment Variable

Set the location of the spreadsheet in an environment variable, otherwise you will be prompted for the location on execution

`export PACKET_RACK_SHEET=<location of spreadsheet>`

## Run the Script

`./rackngo`

## Expected Behavior

RackNGo will find pre-defined devices in an xlsx spreadsheet and add them to or remove them from the API

## License

The source code for the Equinix Metal Console is made available under the terms of the GNU Affero General Public License (GNU AGPLv3). See the LICENSE file for more details.
