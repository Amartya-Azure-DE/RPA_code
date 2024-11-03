def sftp_autodownload(self):
    '''Downloads file automatically from ssh using paramiko'''

    # Define the SFTP server details
    sftp_host = '172.23.34.33'
    sftp_port = 22
    sftp_user = 'nms'
    sftp_pass = 'eci_nms'

    # Create an SSH client
    ssh_client = paramiko.SSHClient()

    # Load the system host keys

    ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    # Connect to the SFTP server
    ssh_client.connect(sftp_host, sftp_port, sftp_user, sftp_pass)

    # Create an SFTP client
    sftp = ssh_client.open_sftp()

    # Get the file name from the remote file path
    today = date.today()
    formatted_date = "{:02d}{:02d}{:02d}".format(today.day, today.month, today.year % 100)

    file_name = "TN_PHY_LO_{}.csv".format(formatted_date)
    # Set the remote file path
    remote_file_path = '/sdh_home/nms/' + file_name

    # Construct the complete local file path
    local_file_path = "D:\\serverFiles\\" + file_name
    print(local_file_path)
    # Download the file from the SFTP server
    sftp.get(remote_file_path, local_file_path)

    # Disconnect from the SFTP server"""
    sftp.close()
    ssh_client.close()