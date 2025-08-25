# Document Library SPFx Solution

This SharePoint Framework solution contains web parts for working with document libraries.

## Web Parts

### 1. Document Library Views Web Part

A web part that displays tiles for each view in a document library. Each tile shows:
- View name
- View type (Standard, Grid, Calendar, etc.)
- Default view indicator
- Hidden view indicator

#### Features:
- **Responsive Design**: Tiles adapt to different screen sizes
- **Interactive**: Click on any tile to open that view in a new tab
- **Visual Indicators**: Clear badges for default and hidden views
- **Error Handling**: Graceful error messages and loading states
- **Configurable**: Set site URL and document library title via properties

#### Configuration:
1. **Site URL**: The URL of the site containing the document library (optional - defaults to current site)
2. **Document Library Title**: The title of the document library (required - no default)
3. **Description**: Optional description text displayed above the tiles

#### Usage:
1. Add the web part to a SharePoint page
2. Configure the properties in the property pane
3. The web part will automatically load and display all views from the specified document library

#### Technical Details:
- Uses SharePoint REST API to fetch views
- Built with React and Fluent UI components
- Supports theme variants (light/dark mode)
- Compatible with SharePoint, Teams, and Office environments

## Development

### Prerequisites
- Node.js (>=16.13.0 <17.0.0 || >=18.17.1 <19.0.0)
- SharePoint Framework development environment

### Building the Solution
```bash
npm install
npm run build
```

### Testing Locally
```bash
npm run serve
```

### Package for Deployment
```bash
npm run package-solution
```

## Deployment

1. Build the solution: `npm run build`
2. Package the solution: `npm run package-solution`
3. Upload the generated `.sppkg` file to your SharePoint App Catalog
4. Add the web part to your pages

## File Structure

```
src/
├── webparts/
│   ├── crestDocument/          # Original web part
│   └── documentLibraryViews/   # New Document Library Views web part
│       ├── components/
│       │   ├── DocumentLibraryViews.tsx
│       │   ├── DocumentLibraryViews.module.scss
│       │   ├── IDocumentLibraryViewsProps.ts
│       │   ├── IDocumentLibraryViewsState.ts
│       │   └── IViewInfo.ts
│       ├── loc/
│       │   ├── en-us.js
│       │   └── mystrings.d.ts
│       ├── DocumentLibraryViewsWebPart.ts
│       ├── DocumentLibraryViewsWebPart.manifest.json
│       └── DocumentLibraryViewsWebPart.module.scss
└── index.ts
```
