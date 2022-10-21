from http import client
import boto3 as boto


# low level client representation
client = boto.client('stepfunctions')

# Getting all execution of ith state machine
response = client.list_executions(
    stateMachineArn='string',
    # statusFilter='RUNNING'|'SUCCEEDED'|'FAILED'|'TIMED_OUT'|'ABORTED',
    # only want succeeded and failed onces
    statusFilter= 'SUCCEEDED'|'FAILED',
    # Just want the first recent execution
    maxResults=1,
    # nextToken='string'
)

_executionArn = response['executions'][0]['executionArn']


# Getting the reference to recent first execution of ith statemachine
_sfLast = client.describe_execution(
    executionArn='excution arn'
)

_status = _sfLast['status']
_startDate = _sfLast['startDate']
_endDate = _sfLast['stopDate']
# {
#     'executionArn': 'string',
#     'stateMachineArn': 'string',
#     'name': 'string',
#     'status': 'RUNNING'|'SUCCEEDED'|'FAILED'|'TIMED_OUT'|'ABORTED',
#     'startDate': datetime(2015, 1, 1),
#     'stopDate': datetime(2015, 1, 1),
#     'input': 'string',
#     'inputDetails': {
#         'included': True|False
#     },
#     'output': 'string',
#     'outputDetails': {
#         'included': True|False
#     },
#     'traceHeader': 'string'
# }