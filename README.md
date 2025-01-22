# Export Active Directory Users to Excel

This script is designed to export a list of users from Active Directory to Excel.

## Features
- **LDAP Integration**: Connects and queries AD servers efficiently.
- **Configurable**: Supports configuration via command-line arguments or a configuration file.
- **Error Handling**: Provides detailed error messages for troubleshooting.
- Search for users using a specified LDAP filter.
- Export the list of users to Excel.

## Installation

1. Ensure Python version 3.7 or higher is installed on your computer.
2. Install the required dependencies using the following command:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Command-Line Arguments
The script accepts the following arguments:

| Argument        | Description                                         | Required | Example                                |
|-----------------|-----------------------------------------------------|----------|----------------------------------------|
| `--server`      | LDAP server address                                | Yes      | `ldap://ad.example.com`               |
| `--user`        | User to connect with                               | Yes      | `user@example.com`                    |
| `--password`    | Password for the user                              | Yes      | `password123`                         |
| `--base`        | Base DN for searching                              | Yes      | `DC=example,DC=com`                   |
| `--filter`      | LDAP filter for searching users                    | No       | `(objectClass=user)`                  |
| `--output`      | Path to output file xlsx                  | Yes      | `output.xlsx`                         |
| `--config`      | Path to a configuration file (optional alternative)| No       | `config.ini`                          |

### Example Usage
#### Command-Line Execution
```bash
python ad_script.py \
    --server ldap://ad.example.com \
    --user user@example.com \
    --password password123 \
    --base DC=example,DC=com \
    --filter "(objectClass=user)" \
    --output users.xlsx
```

After execution, the script will create a file `users.xlsx` containing user data.


#### Using a Configuration File
Create a `config.ini` file:
```ini
[AD]
server=ldap://ad.example.com
user=user@example.com
password=password123
base=DC=example,DC=com
filter=(objectClass=user)
```

Run the script:
```bash
python ad_script.py --config config.ini --output users.xlsx
```

---

## Troubleshooting

1. **Connection Issues**:
   - Ensure the server address is correct and reachable.
   - Verify that the provided credentials have sufficient permissions.

2. **Invalid Filter**:
   - Confirm the LDAP filter syntax is correct.

3. **Output Errors**:
   - Ensure the specified output path is writable.

---

## Contributing
Contributions are welcome! Please follow these steps:
1. Fork the repository.
2. Create a feature branch.
3. Commit changes and push to your fork.
4. Create a pull request.

---

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
