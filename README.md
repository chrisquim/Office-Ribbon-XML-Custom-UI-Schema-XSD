# Microsoft Office Ribbon XML Custom UI Schema (XSD)

This repository provides the XML Schema Definition (XSD) files for validating Microsoft Office Ribbon XML customizations for both Office 2007 and Office 2010.  These XSD files can be invaluable for developers creating custom Ribbon UI elements for their Office applications.

## Table of Contents

- [Introduction](#introduction)
- [Schema Versions](#schema-versions)
  - [Office 2007](#office-2007)
  - [Office 2010](#office-2010)
- [Validation with Notepad++](#validation-with-notepad)
  - [Prerequisites](#prerequisites)
  - [Validation Steps](#validation-steps)
- [Example Usage](#example-usage)
- [Resources](#resources)

## Introduction

The Ribbon UI, introduced in Microsoft Office 2007, allows developers to significantly enhance the user experience by creating custom toolbars, menus, and other interactive elements within Office applications.  Customizing the Ribbon UI involves creating XML markup that defines the structure and behavior of your custom elements.

To ensure that your Ribbon XML is well-formed and adheres to the correct syntax, it is essential to validate it against the appropriate XSD schema. Validation helps identify errors early in the development process, saving you time and effort.

## Schema Versions

### Office 2007

- **Filename:** `Office 2007 CustomUI Schema.xsd`
- **Source Link:** [https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/5f3e35d6-70d6-47ee-9e11-f5499559f93a](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/5f3e35d6-70d6-47ee-9e11-f5499559f93a)

### Office 2010

- **Filename:** `Office 2010 CustomUI Schema.xsd`
- **Source Link:** [https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui2/daabc2cd-b3f5-4d32-9099-95d705f70f35](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui2/daabc2cd-b3f5-4d32-9099-95d705f70f35)

**Note:** Be sure to use the correct XSD version that corresponds to the target version of Microsoft Office for your Ribbon customization.

## Validation with Notepad++

Notepad++ is a popular, free, and open-source code editor that, when combined with the XML Tools plugin, provides a convenient way to validate your Ribbon XML.

### Prerequisites

1. **Notepad++:**  Download and install Notepad++ from the official website: [https://notepad-plus-plus.org/](https://notepad-plus-plus.org/)
2. **XML Tools Plugin:**
   - Open Notepad++.
   - Go to **Plugins** > **Plugins Admin...**
   - Search for "XML Tools" and install the plugin. 
   - Restart Notepad++ after installation.

### Validation Steps

1. **Open your Ribbon XML file in Notepad++.**
2. **Associate the XSD Schema:**
   - Go to **Plugins** > **XML Tools** > **Validate Now**. 
   - In the "Validate" dialog box, click on the **...** button next to the "DTD/Schema" field.
   - Locate and select the appropriate XSD file (either `customUI.xsd` or `customUI14.xsd`) from this repository.
   - Click **Open** to associate the schema.
3. **Validate Your XML:**
   - Click **OK** in the "Validate" dialog box.
   - Notepad++ will validate your XML against the chosen schema and display any errors or warnings in the XML Tools output window.

## Example Usage

Here is a basic example of how you can use these XSD files to validate your Ribbon XML:

**Example Ribbon XML (customUI.xml):**

```xml
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="customTab" label="My Custom Tab">
        <group id="customGroup" label="My Group">
          <button id="myButton" label="Click Me" size="large"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

**Validation:**

1. Follow the steps in the "Validation with Notepad++" section above.
2. If your XML is valid, you will see a success message in the XML Tools output window.

## Resources

- **Microsoft Office Developer Documentation:**  [https://docs.microsoft.com/en-us/office/](https://docs.microsoft.com/en-us/office/)
- **RibbonX and Custom UI Editor:** This tool can help you visually design your Ribbon UI: [https://github.com/fernandreu/RibbonXEditor](https://github.com/fernandreu/RibbonXEditor) 
