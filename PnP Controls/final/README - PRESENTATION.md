## pnp-controls

Steps to follow during the demo

### Before the demo

Ensure you start from the "start" folder under the PnP Controls directory
The start folder already contains a React web part with the required additional modules installed:
-PnPJS
-PnP reusable controls
-PnP reusable property-pane controls

1) Install PnPJS
npm install @pnp/logging @pnp/common @pnp/odata @pnp/sp @pnp/graph --save

2) Install PnP reusable controls
npm install @pnp/spfx-controls-react --save --save-exact

Configure resource file by adding the below to config/config.json
"ControlStrings": "node_modules/@pnp/spfx-controls-react/lib/loc/{locale}.js"

3) Install PnP reusable property-pane controls
npm install @pnp/spfx-property-controls --save --save-exact

Configure resource file by adding the below to config/config.json
"PropertyControlStrings": "node_modules/@pnp/spfx-property-controls/lib/loc/{locale}.js"
