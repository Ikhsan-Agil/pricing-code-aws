import pandas as pd
import re
import json
import boto3
import os
import datetime
from dotenv import load_dotenv

# load .env file
load_dotenv()

def get_aws_session():
    try:
        # Try creating a session with the specified SSO profile
        session = boto3.Session(profile_name=os.getenv("AWS_PROFILE_NAME"), region_name="us-east-1")
        print(f"Successfully loaded AWS session using profile: {os.getenv("AWS_PROFILE_NAME")}")
        return session
    except Exception as e:
        print(f"Warning: Could not load profile '{os.getenv("AWS_PROFILE_NAME")}' ({e}). Falling back to default session.")
        return boto3.Session(region_name="us-east-1")

def fetch_instance_type_from_aws(ec2_client, cpu, memory_gb, production):
    """
    Step 4: Dynamically search/fetch available EC2 instance types in us-east-1
    using AWS SDK (describe_instance_types) based on CPU, memory, and production flag.
    """
    is_prod = str(production).strip().lower() == 'yes'
    
    if is_prod:
        ratio = memory_gb / cpu if cpu > 0 else 2
        if ratio > 4:
            allowed_prefixes = ['r6i.', 'r5.'] # Memory optimized
        elif ratio < 2:
            allowed_prefixes = ['c6i.', 'c5.'] # Compute optimized
        else:
            allowed_prefixes = ['m6i.', 'm5.'] # General purpose
    else:
        allowed_prefixes = ['t3.', 't4g.'] # Non-production / Burstable
        
    memory_mib = int(memory_gb * 1024)
    
    try:
        paginator = ec2_client.get_paginator('describe_instance_types')
        page_iterator = paginator.paginate(
            Filters=[
                {'Name': 'vcpu-info.default-vcpus', 'Values': [str(int(cpu))]}
            ]
        )
        
        matching_instances = []
        for page in page_iterator:
            for itype in page['InstanceTypes']:
                instance_name = type_name = itype['InstanceType']
                
                if not any(type_name.startswith(p) for p in allowed_prefixes):
                    continue
                    
                if 'MemoryInfo' in itype:
                    inst_mem_mib = itype['MemoryInfo']['SizeInMiB']
                    if inst_mem_mib >= memory_mib:
                        matching_instances.append((type_name, inst_mem_mib))
                        
        if matching_instances:
            matching_instances.sort(key=lambda x: (x[1], x[0]))
            return matching_instances[0][0]
            
    except Exception as e:
        print(f"Warning: Error querying EC2 instance types from AWS SDK: {e}")
        
    if is_prod:
        return "m6i.large" if cpu <= 2 else "m6i.xlarge"
    else:
        return "t3.medium"

def fetch_ec2_price_from_aws(pricing_client, instance_type):
    """Step 6: Fetch live EC2 hourly price using AWS Price List API."""
    try:
        response = pricing_client.get_products(
            ServiceCode='AmazonEC2',
            Filters=[
                {'Type': 'TERM_MATCH', 'Field': 'instanceType', 'Value': instance_type},
                {'Type': 'TERM_MATCH', 'Field': 'location', 'Value': 'US East (N. Virginia)'},
                {'Type': 'TERM_MATCH', 'Field': 'tenancy', 'Value': 'shared'},
                {'Type': 'TERM_MATCH', 'Field': 'preInstalledSw', 'Value': 'NA'},
                {'Type': 'TERM_MATCH', 'Field': 'operation', 'Value': 'RunInstances'}
            ],
            MaxResults=1
        )
        if not response.get('PriceList'):
            return 0.10
        
        price_item = json.loads(response['PriceList'][0])
        terms = price_item.get('terms', {}).get('OnDemand', {})
        term_id = list(terms.keys())[0]
        price_dimensions = terms[term_id].get('priceDimensions', {})
        dim_id = list(price_dimensions.keys())[0]
        
        return float(list(price_dimensions[dim_id]['pricePerUnit'].values())[0])
    except Exception as e:
        print(f"Warning: Could not fetch live price for {instance_type} ({e}). Using fallback rate.")
        return 0.10

def fetch_storage_price_from_aws(pricing_client, volume_api_name):
    """Step 5 & 6: Fetch EBS per-GB storage price using AWS Price List API."""
    try:
        response = pricing_client.get_products(
            ServiceCode='AmazonEC2',
            Filters=[
                {'Type': 'TERM_MATCH', 'Field': 'volumeApiName', 'Value': volume_api_name},
                {'Type': 'TERM_MATCH', 'Field': 'location', 'Value': 'US East (N. Virginia)'}
            ],
            MaxResults=1
        )
        if not response.get('PriceList'):
            return 0.125 if volume_api_name == 'io2' else 0.08
        
        price_item = json.loads(response['PriceList'][0])
        terms = price_item.get('terms', {}).get('OnDemand', {})
        term_id = list(terms.keys())[0]
        price_dimensions = terms[term_id].get('priceDimensions', {})
        
        for dim_id, dim_val in price_dimensions.items():
            price_per_unit = dim_val.get('pricePerUnit', {})
            if 'USD' in price_per_unit:
                return float(price_per_unit['USD'])
        return 0.08
    except Exception as e:
        return 0.125 if volume_api_name == 'io2' else 0.08

def parse_disk_size_total(size_str):
    """Helper to calculate total numeric GB from strings, supporting '+' symbols (e.g., '25GB+70GB')."""
    if pd.isna(size_str) or str(size_str).strip() in ["", "0GB", "nan"]:
        return 0
    matches = re.findall(r'(\d+)', str(size_str))
    return sum(int(m) for m in matches) if matches else 0

def main():
    date = datetime.datetime.now().strftime("%Y-%m-%dat%H-%M-%SWIB")
    input_file = os.path.join("input","pricing-input.xlsx")
    output_name = "aws-final-pricing-output-" + date + ".xlsx"
    output_file = os.path.join("output", output_name)
    
    # Step 2: Initialize AWS APIs using the explicit SSO profile session
    print("Initializing AWS SDK session and clients for us-east-1...")
    aws_session = get_aws_session()
    ec2_client = aws_session.client('ec2', region_name='us-east-1')
    pricing_client = aws_session.client('pricing', region_name='us-east-1')
    
    # Step 3: Read all the Excel input
    print(f"Reading input file: {input_file}...")
    raw_df = pd.read_excel(input_file, header=None)
    
    # Cache storage pricing rates
    gp3_rate = fetch_storage_price_from_aws(pricing_client, 'gp3')
    io2_rate = fetch_storage_price_from_aws(pricing_client, 'io2')
    
    processed_rows = []
    
    # Iterate through server data rows starting from row index 2
    for i in range(2, len(raw_df)):
        row = raw_df.iloc[i]
        if pd.isna(row[1]) or pd.isna(row[0]):
            continue
            
        no_id = row[0]
        server_name = row[1]
        cpu = float(row[2])
        memory = float(row[3])
        sys_disk_raw = row[4] if not pd.isna(row[4]) else "0GB"
        data_disk_raw = row[5] if not pd.isna(row[5]) else "0GB"
        production = row[6]
        high_io = row[7]
        
        # Step 4: Search available EC2 instance type via AWS SDK
        print(f"Querying AWS SDK for server '{server_name}' (vCPU: {cpu}, Memory: {memory}GB, Production: {production})...")
        instance_type = fetch_instance_type_from_aws(ec2_client, cpu, memory, production)
        print(f" -> Selected Instance Type: {instance_type}")
        
        # Step 5: Examine disk data (High I/O -> io2, else -> gp3) and size
        is_high_io = str(high_io).strip().lower() == 'yes'
        disk_type_label = "io2" if is_high_io else "gp3"
        storage_unit_rate = io2_rate if is_high_io else gp3_rate
        
        sys_gb = parse_disk_size_total(sys_disk_raw)
        data_gb = parse_disk_size_total(data_disk_raw)
        
        sys_disk_formatted = f"{sys_disk_raw} ({disk_type_label})" if sys_gb > 0 else "0GB"
        data_disk_formatted = f"{data_disk_raw} ({disk_type_label})" if data_gb > 0 else "0GB"
        
        total_disk_gb = sys_gb + data_gb
        
        # Step 6: Count the price from all fetched specifications
        hourly_compute_rate = fetch_ec2_price_from_aws(pricing_client, instance_type)
        
        monthly_compute_cost = hourly_compute_rate * 730 # 730 hours in a standard month
        monthly_storage_cost = total_disk_gb * storage_unit_rate
        total_monthly_price = round(monthly_compute_cost + monthly_storage_cost, 2)
        
        # Step 7: Create the Excel output rows including the No. column and separated disks
        processed_rows.append({
            "No.": no_id,
            "Instance Name": server_name,
            "Instance Type": instance_type,
            "CPU Cores": cpu,
            "Memory in GB": memory,
            "System Disk": sys_disk_formatted,
            "Data Disk": data_disk_formatted,
            "Price ($/month)": total_monthly_price
        })
        
    output_df = pd.DataFrame(processed_rows)
    output_df.to_excel(output_file, index=False)
    print(f"\nEstimation complete! Output successfully saved to: {output_file}")
    print(output_df)

if __name__ == "__main__":
    main()