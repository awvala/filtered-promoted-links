## filtered-promoted-links

This is a custom Promoted Links web part that enables a filtered, cascading dropdown field in the property pane panel.  The hover description panel also includes an Owner field property to identify a user or contact for each promoted link.  This web part is compatible with both the Modern and Classic SharePoint experience.  

### Features
* Wrapped promoted link tiles.
* Cascading dropdown options between a Promoted Links list and a filter column.
* Placeholder and Spinner elements to handle unconfigured web parts and empty lists.
* Links configured to open in current window or in new tab based on Launch Behavior selection in the Promoted Links list. 
  * "Dialog" selections will open in a new tab. 

### Required List Settings
* Promoted Link list
    * Choice field named "Filter"
    * People Picker field named "Owner"

### SPFx Add-ons and tools
* PropertyPaneDropdown, IPropertyPaneDropdownOption
* spfx-controls-react
    * Placeholder control
* office-ui-fabric-react
    * Spinner, SpinnerSize
    * Image, IImageProps, ImageFit
* SPHttpClient

### Demos

#### Adding Web Part to a Modern Page
![Full Demo of the Modern Promoted Links](/src/assets/placeholder)

---

#### Adding Web Part to a Classic Page
![Full Demo of the Modern Promoted Links in a classic page](/src/assets/placeholder)

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```
