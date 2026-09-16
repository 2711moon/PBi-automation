import getpass, sys, keyring

SERVICE_NAME = 'PBI_Automation_Kisna'

def main():
    print()
    print('=' * 60)
    print('  Power BI Credential Setup')
    print('  Saved to Windows Credential Manager - never to any file.')
    print('=' * 60)
    print()
    username = input('  Power BI admin email    : ').strip()
    if not username:
        print('ERROR: Email cannot be empty.')
        sys.exit(1)
    password = getpass.getpass('  Power BI admin password : ')
    if not password:
        print('ERROR: Password cannot be empty.')
        sys.exit(1)
    keyring.set_password(SERVICE_NAME, 'username', username)
    keyring.set_password(SERVICE_NAME, 'password', password)
    print()
    print('  Credentials saved. Verify at:')
    print('  Control Panel -> Credential Manager -> Windows Credentials')
    print(f'  Service name: {SERVICE_NAME}')
    print()
    print('  Next step: fill in PBI_REPORT_URL in config.py, then run main.py')
    print()

if __name__ == '__main__':
    main()
