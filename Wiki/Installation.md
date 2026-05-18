# Installation

To get started with oPenEfficiency, you'll need a Windows environment with PowerPoint installed.

## Prerequisites
- **Operating System:** Windows 10 or Windows 11
- **Application:** Microsoft PowerPoint 2016 or later
- **Framework:** .NET Framework 4.8
- **Runtime:** [Visual Studio 2022 Tools for Office Runtime](https://aka.ms/vstor_100_150)

## Manual Installation (Development)

If you are a developer or want to build the add-in from source, follow these steps:

1. **Clone the repository:**
   Clone or download the project repository to your local machine.

2. **Open the Solution:**
   Open the `oPenEfficiency.slnx` or `oPenEfficiency.csproj` file in Visual Studio 2022.

3. **Install Workloads:**
   Ensure that the **"Office/SharePoint development"** workload is installed in your Visual Studio environment.

4. **Build the Solution:**
   Build the solution in **Release** configuration. This will compile the Add-in and automatically register it on your machine.

5. **Restart PowerPoint:**
   Close any running instances of PowerPoint and restart the application. You should now see the new "oPen Efficiency" ribbon tab.