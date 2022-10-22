import boto3



client = boto3.client('s3')
bucket = lambda_event["bucket"]

archive_file_path = "veeva_test/outbound_may_be/archive"
file_name = "itachi.xlsx"

# uploading loacl archive file to S3
client.upload_file(archive_file_path + '/' + file_name, bucket, )
    