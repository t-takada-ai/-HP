// Render react-icons (Feather) to base64 PNGs for embedding in pptx
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const Fi = require("react-icons/fi");
const sharp = require("sharp");

const NAMES = [
  "FiUsers", "FiTrendingDown", "FiTrendingUp", "FiAlertTriangle", "FiFileText",
  "FiClock", "FiDollarSign", "FiSearch", "FiEdit3", "FiRepeat", "FiCheckSquare",
  "FiBriefcase", "FiSettings", "FiTarget", "FiMessageCircle", "FiUserCheck",
  "FiShield", "FiZap", "FiLayers", "FiMap", "FiCalendar", "FiAward", "FiPlay",
  "FiBookOpen", "FiPieChart", "FiBarChart2", "FiArrowRight", "FiPhone",
  "FiUserX", "FiSlash", "FiThumbsUp", "FiFlag", "FiCpu", "FiClipboard",
  "FiFilter", "FiGrid", "FiActivity", "FiHelpCircle", "FiLock", "FiMinimize2",
  "FiMonitor", "FiSend", "FiLink", "FiSmartphone", "FiTerminal", "FiFolder",
  "FiMail", "FiDatabase", "FiShoppingCart", "FiPhoneCall", "FiCloud", "FiMousePointer",
];

async function renderAll(colors) {
  const map = {};
  for (const name of NAMES) {
    const Comp = Fi[name];
    if (!Comp) throw new Error("missing icon " + name);
    for (const [key, hex] of Object.entries(colors)) {
      const svg = ReactDOMServer.renderToStaticMarkup(
        React.createElement(Comp, { color: "#" + hex, size: 256, strokeWidth: 2 })
      );
      const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
      map[name + ":" + key] = "image/png;base64," + buf.toString("base64");
    }
  }
  return map;
}

module.exports = { renderAll, NAMES };
