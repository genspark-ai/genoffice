// Implementation lives in the engine (parse-time TOC numbering reuses it);
// this module keeps the historical import path for the editor.
export {
  substituteBullet,
  markerFallbackFace,
  SEGOE_UI_SYMBOL_RE,
  computeListMarkerInfos,
  computeListMarkers,
  formatNumber,
  markerTabAdvance,
  type ListItemRef,
  type ListMarkerInfo,
} from '@genoffice/docx-engine'
