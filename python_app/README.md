# Microsoft Graph API Explorer - Dynamic Filters

A comprehensive Python GUI application for exploring Microsoft Graph APIs with advanced dynamic filter builders, MSAL authentication, and nested folder browsing capabilities.

## Features

### 🔧 **Dynamic Filter Builders**

- **Boolean Filters**: True/false dropdown selections for properties like `isRead`, `hasAttachments`
- **DateTime Filters**: Custom date/time input with ISO format support (`YYYY-MM-DDTHH:MM:SSZ`)
- **Number Filters**: Numeric input with validation and range constraints (`$top`, `$skip`)
- **Text Filters**: Free-text search for content, subjects, and custom queries
- **Email Filters**: Specific email address filtering for sender/recipient queries
- **Select Filters**: Dropdown options for predefined values (importance levels, etc.)
- **Multi-select Filters**: Checkbox selections for field choosing (`$select` parameters)
- **Compound Filters**: Complex filters with multiple input fields (`$orderBy` with direction)
- **Static Filters**: Pre-configured filters requiring no user input (`$expand`)

### 🔐 **MSAL Authentication Integration**

- **Interactive Authentication**: Full OAuth flow with Microsoft Graph
- **Token Management**: Automatic token acquisition and refresh
- **Real-time Status**: Visual authentication status indicators
- **Live API Calls**: Execute queries directly against Microsoft Graph
- **Error Handling**: Comprehensive authentication error management

### 📁 **Advanced Folder Management**

- **Nested Folder Support**: Recursive folder loading up to 10 levels deep
- **Hierarchical Display**: Visual folder tree with indentation and icons
- **Folder Statistics**: Total and unread item counts for each folder
- **Path Preservation**: Full hierarchy paths (e.g., `Inbox/Projects/2025/January`)
- **Folder Selection**: Target specific folders for message queries
- **Interactive Browser**: Pop-up folder browser with search and selection

### 🎯 **Smart Query Building**

- **Real-time Preview**: Live preview of filter construction as you type
- **Filter Combination**: Add/remove multiple filters with visual feedback
- **Template System**: Variable substitution in filter templates
- **URL Generation**: Complete Microsoft Graph URLs with proper parameter encoding
- **Clipboard Integration**: One-click copying of generated URLs
- **Query Execution**: Direct API calls with formatted results

### 🎨 **Modern Dark Theme UI**

- **Professional Design**: Microsoft-inspired dark theme throughout
- **Responsive Layout**: Equal-split panels for endpoints and filter builders
- **Visual Feedback**: Button state changes and status indicators
- **Scrollable Content**: Handle large datasets with smooth scrolling
- **Color-coded Elements**: Method badges, category tags, and status indicators

### 📊 **Comprehensive Results Display**

- **Formatted Results**: Structured display of API responses
- **Message Details**: Subject, sender, date, read status, attachments
- **Folder Information**: Display names, item counts, hierarchy
- **Error Reporting**: Detailed error messages with status codes
- **Result Limitations**: Smart truncation for large datasets

## Installation

### Prerequisites

- Python 3.8 or higher
- pip (Python package installer)
- Internet connection for MSAL authentication

### Setup

1. **Clone or download the project:**

   ```bash
   git clone <repository-url>
   cd python_app
   ```

2. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

   Or install manually:

   ```bash
   pip install pyperclip msal requests beautifulsoup4
   ```

3. **Configure Azure App Registration (Optional for MSAL):**

   - Copy `config_sample.cfg` to `config.cfg`
   - Update with your Azure app credentials:

   ```ini
   [azure]
   clientId = your-azure-app-client-id
   tenantId = your-azure-tenant-id
   ```

4. **Run the application:**

   ```bash
   python main_dynamic_filters.py
   ```

   Or use the launcher:

   ```bash
   python run.py
   ```

## Usage

### 1. **Launch and Authenticate**

```bash
python main_dynamic_filters.py
```

- Click "🔐 Authenticate" to sign in with Microsoft Graph (optional)
- Authentication enables live API calls and folder browsing

### 2. **Browse API Endpoints**

- **Search**: Type keywords to find specific endpoints
- **Filter by Category**: Choose from Mail, Calendar, Users, etc.
- **Filter by Version**: Select v1.0 or beta APIs
- **Method Filter**: Filter by GET, POST, PUT, DELETE
- **Click Endpoints**: Select an endpoint to build custom filters

### 3. **Build Dynamic Filters**

- **Select Endpoint**: Click any endpoint card to open the filter builder
- **Choose Filters**: Each endpoint shows available filter types
- **Configure Values**:
  - Boolean: Select true/false from dropdown
  - Text: Enter search terms or email addresses
  - Number: Input numeric values with validation
  - DateTime: Enter ISO format dates
  - Multi-select: Check/uncheck multiple options
- **Real-time Preview**: See filter construction as you type
- **Add Filters**: Click "✓ Add Filter" to include in query
- **Remove Filters**: Click "✗ Remove" to exclude from query

### 4. **Manage Folders (Mail Endpoints)**

- **Browse Folders**: Click "🗂️ Browse Folders" after authentication
- **Navigate Hierarchy**: See nested folders with visual indentation
- **View Statistics**: Total/unread counts for each folder
- **Select Folder**: Choose target folder for message queries
- **Folder Filtering**: Automatically applies folder context to queries

### 5. **Execute and Copy Queries**

- **Execute Query**: Click "🚀 Execute Query with Custom Filters"
- **View Results**: See formatted API responses in popup window
- **Copy URLs**: Click "📋 Copy" to copy individual filters
- **Clear Filters**: Click "🗑️ Clear Selected Filters" to reset
- **Export Results**: Copy complete URLs to clipboard

## Supported Microsoft Graph Endpoints

### **Mail Operations**

- **List Messages**: Get user messages with extensive filtering
  - Read status, date ranges, attachments, importance
  - Sender filtering, content search, field selection
  - Custom sorting and result limiting
- **List Mail Folders**: Browse folder hierarchy
  - Folder name filtering, field selection
  - Child folder expansion, sorting options

### **Filter Types by Endpoint**

| Filter Type      | Description              | Example Usage                                    |
| ---------------- | ------------------------ | ------------------------------------------------ |
| **Boolean**      | True/false selections    | `isRead eq true`, `hasAttachments eq false`      |
| **DateTime**     | ISO date/time filtering  | `receivedDateTime ge 2025-01-15T00:00:00Z`       |
| **Number**       | Numeric with validation  | `$top=50`, `$skip=25`                            |
| **Text**         | Free-text search         | `$search="project update"`                       |
| **Email**        | Email address filtering  | `from/emailAddress/address eq 'user@domain.com'` |
| **Select**       | Predefined options       | `importance eq 'high'`                           |
| **Multi-select** | Multiple field selection | `$select=subject,from,receivedDateTime`          |
| **Compound**     | Multiple related inputs  | `$orderBy=receivedDateTime desc`                 |
| **Static**       | No input required        | `$expand=childFolders`                           |

## MSAL Integration Details

### **Authentication Flow**

1. Interactive OAuth login via browser
2. Token acquisition and storage
3. Automatic token refresh
4. Secure API calls with Bearer tokens

### **Required Permissions**

- `Mail.Read`: Read user mail messages
- `Mail.ReadWrite`: Read and modify user mail

### **Configuration**

Create `config.cfg` with your Azure app details:

```ini
[azure]
clientId = 12345678-1234-1234-1234-123456789012
tenantId = 87654321-4321-4321-4321-210987654321
```

## Dependencies

### **Core Requirements**

- **tkinter**: GUI framework (included with Python)
- **pyperclip**: Clipboard operations for URL copying

### **MSAL Integration**

- **msal**: Microsoft Authentication Library
- **requests**: HTTP client for API calls
- **beautifulsoup4**: HTML parsing for enhanced functionality

### **Optional**

- **python-dateutil**: Enhanced date/time handling
- **configparser**: Configuration file management

## Troubleshooting

### **Missing tkinter**

- **Ubuntu/Debian**: `sudo apt-get install python3-tk`
- **CentOS/RHEL**: `sudo yum install tkinter`
- **macOS**: Included with Python from python.org
- **Windows**: Included with Python

### **MSAL Authentication Issues**

- Verify Azure app registration settings
- Check client ID and tenant ID in config.cfg
- Ensure proper redirect URI configuration
- Review permission scopes in Azure portal

### **API Call Failures**

- Check internet connectivity
- Verify authentication status
- Review Microsoft Graph service status
- Check rate limiting and throttling

## File Structure

```
python_app/
├── main.py    # Main application with all features
├── run.py                     # Launcher script with dependency checking
├── requirements.txt           # Python dependencies
├── config_sample.cfg          # Sample Azure configuration
├── README.md                  # This documentation
├── setup.py                   # Package setup for distribution
└── documentation.txt      # Detailed API documentation
```

## Development

### **Building Distribution**

```bash
python setup.py sdist bdist_wheel
```

### **Code Quality**

```bash
pip install black flake8 mypy
black main_dynamic_filters.py
flake8 main_dynamic_filters.py
mypy main_dynamic_filters.py
```

### **Testing**

```bash
pip install pytest pytest-cov
pytest tests/
```

## Advanced Features

### **Filter Persistence**

- Selected filters remain active until manually cleared
- Button states reflect current filter selection
- Real-time preview updates with each change

### **Error Handling**

- Comprehensive validation for all input types
- Network error recovery and reporting
- Authentication state management

### **Performance Optimization**

- Lazy loading of nested folders
- Efficient UI updates and rendering
- Smart caching of API responses

## License

This project is open source and available under the MIT License.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## Support

For issues and questions:

- Check the troubleshooting section
- Review Microsoft Graph documentation
- Open an issue on the project repository
