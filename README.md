# SolidWorks 2022 C# Add-in Template

A complete, functional SolidWorks add-in template written in C# that demonstrates common add-in functionality and provides a solid foundation for developing custom SolidWorks extensions.

## Features

- **Command Toolbar**: Custom toolbar with sample buttons
- **Property Manager Page**: Interactive UI with various control types (textboxes, checkboxes, radio buttons, listboxes, etc.)
- **Event Handling**: Framework for handling SolidWorks and document events
- **COM Registration**: Automatic registration for SolidWorks integration
- **Sample Functionality**: 
  - Create Cube: Demonstrates basic part creation and feature modeling
  - Property Manager: Shows how to create interactive user interfaces

## Requirements

- **SolidWorks 2022** (configured for SolidWorks 2022, can be adapted for other versions)
- **.NET Framework 4.8.1**
- **Visual Studio 2019/2022** (must be run as Administrator for COM registration)
- **Windows** (SolidWorks requirement)

## Getting Started

### 1. Clone the Repository
```bash
git clone [your-repo-url]
cd swcsharpaddin
```

### 2. Open in Visual Studio
- **Important**: Run Visual Studio as Administrator (required for COM registration)
- Open `swcsharpaddin.csproj` or the solution file

### 3. Build and Run
- Press `F5` to build and launch SolidWorks with the add-in loaded
- Look for the "C# Addin" toolbar in SolidWorks

### 4. Test the Add-in
- **Create Cube**: Click to create a simple extruded cube in a new part
- **Show PMP**: Click to display the sample Property Manager Page with various controls

## Project Structure

```
swcsharpaddin/
??? SwAddin.cs              # Main add-in class implementing ISwAddin
??? UserPMPage.cs           # Property Manager Page implementation
??? PMPHandler.cs           # Property Manager Page event handler
??? EventHandling.cs        # Document and application event handling
??? AssemblyInfo.cs         # Assembly metadata and COM visibility
??? swcsharpaddin.csproj    # Project file with SolidWorks API references
??? *.bmp                   # Toolbar icon resources
```

## Key Classes

- **SwAddin**: Main add-in entry point, handles connection to SolidWorks
- **UserPMPage**: Creates and manages the Property Manager Page UI
- **PMPHandler**: Handles events from Property Manager Page controls
- **DocumentEventHandler**: Base class for handling document-specific events

## Customization

### Adding New Commands
1. Add command IDs to the constants in `SwAddin.cs`
2. Create command items in the `AddCommandMgr()` method
3. Implement callback methods for your new commands

### Modifying the Property Manager Page
1. Edit `UserPMPage.cs` to add/remove controls
2. Update control IDs and implement event handling in `PMPHandler.cs`

### Adding Event Handlers
1. Extend the event handling classes in `EventHandling.cs`
2. Subscribe to additional SolidWorks events as needed

## Troubleshooting

### Common Issues
- **Add-in not loading**: Ensure Visual Studio is running as Administrator
- **Wrong SolidWorks version launching**: Update paths in project file to point to correct SolidWorks installation
- **Build errors**: Verify SolidWorks API references point to correct installation directory

### Build Requirements
- The project must be built with Administrator privileges for COM registration
- SolidWorks API assemblies must be available in the specified paths

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is provided as-is for educational and development purposes. Please ensure compliance with SolidWorks API licensing terms when using in commercial applications.

## Acknowledgments

- Built using the SolidWorks API
- Template structure based on SolidWorks add-in development best practices
- Targets .NET Framework 4.8.1 for optimal SolidWorks 2022 compatibility