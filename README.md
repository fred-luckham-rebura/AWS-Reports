# AWS-Report-roles-and-entities
A script that uses name profiles in your AWS credentials file to create an Excel workbook of all the roles and their respective trusted entities in the accounts provided. 

## Usage
You must have named (non default) acccount crednentials in your AWS credentials file for this to work. It requires a key with RO access. To use just run the script and the output (excel workbooks) will be created in the same directory. The output can be merged using the openpyxl merger script here: https://github.com/Fred-Luckham/openpyxl-workbook-merger

## Requiremnts
- Boto3
- openpyxl
