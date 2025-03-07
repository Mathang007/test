import boto3
from openpyxl import Workbook
from datetime import datetime, timedelta, timezone

def get_logs_for_multiple_users_to_excel(usernames, aws_region, output_file):
    """
    Fetches logs for multiple users from AWS CloudTrail and writes them to an Excel file with separate sheets.
    
    :param usernames: List of usernames to filter logs.
    :param aws_region: The AWS region to filter logs.
    :param output_file: Path to the Excel file to save logs.
    """
    try:
        # Calculate yesterday's date range (UTC timezone)
        end_time = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)  # Midnight today (UTC)
        start_time = end_time - timedelta(days=1)  # Midnight yesterday (UTC)

        # Initialize CloudTrail client for the specified region
        client = boto3.client('cloudtrail', region_name=aws_region)

        # Create an Excel workbook
        workbook = Workbook()

        for idx, username in enumerate(usernames):
            print(f"Processing logs for user: {username}")
            # Fetch logs using lookup attributes
            logs = []
            paginator = client.get_paginator('lookup_events')
            page_iterator = paginator.paginate(
                LookupAttributes=[{'AttributeKey': 'Username', 'AttributeValue': username}],
                StartTime=start_time,
                EndTime=end_time
            )

            for page in page_iterator:
                logs.extend(page['Events'])

            # Add a new sheet for each user
            sheet_name = username[:30] if len(username) > 30 else username  # Limit sheet name to 30 characters
            if idx == 0:
                sheet = workbook.active
                sheet.title = sheet_name
            else:
                sheet = workbook.create_sheet(title=sheet_name)

            # Write headers to the sheet
            sheet.append([
                'Username', 
                'Event Time', 
                'Event Source', 
                'Event Name', 
                'AWS Region', 
                'Source IP address', 
                'Resources', 
                'Read-only', 
                'Event Type', 
                'Event Category'
            ])

            # Write event details
            for event in logs:
                # Convert timezone-aware datetime to naive datetime
                event_time = event.get('EventTime', 'N/A')
                if event_time != 'N/A':
                    event_time = event_time.replace(tzinfo=None)

                resources = ', '.join(
                    [resource.get('ResourceName', 'N/A') for resource in event.get('Resources', [])]
                ) if event.get('Resources') else 'N/A'

                sheet.append([
                    username,
                    event_time,
                    event.get('EventSource', 'N/A'),
                    event.get('EventName', 'N/A'),
                    event.get('AwsRegion', 'ap-south-1'),
                    event.get('SourceIPaddress', 'N/A'),
                    event.get('resources', 'N/A'),
                    event.get('ReadOnly', 'N/A'),
                    event.get('EventType', 'N/A'),
                    event.get('EventCategory', 'N/A')
                ])
        
        # Save the Excel file
        workbook.save(output_file)
        print(f"Logs for all users saved to {output_file}")

    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage
if __name__ == "__main__":
    usernames = ["Ajith_datautics", "sandru_datautics", "sarath_datautics"]  # Replace with your list of usernames
    aws_region = "ap-south-1"  # Replace with your AWS region
    output_file = f"combined_user_logs_{(datetime.now(timezone.utc) - timedelta(days=1)).strftime('%Y-%m-%d')}.xlsx"

    get_logs_for_multiple_users_to_excel(usernames, aws_region, output_file)

