import { strFromU8 } from "fflate";

import {
  CHART_NS,
} from "./chart-series";

export function normalizeArchivePath(path: string) {
  return path.replace(/^\/+/, "").replace(/\\/g, "/");
}

export function dirname(path: string) {
  const normalized = normalizeArchivePath(path);
  const index = normalized.lastIndexOf("/");
  return index >= 0 ? normalized.slice(0, index) : "";
}

export function resolveRelationshipPath(basePath: string, target: string) {
  if (!target) {
    return "";
  }

  const normalizedTarget = target.replace(/\\/g, "/");
  if (normalizedTarget.startsWith("/")) {
    return normalizeArchivePath(normalizedTarget);
  }
  const normalizedBasePath = normalizeArchivePath(basePath);
  let baseDirectory = dirname(normalizedBasePath);
  if (normalizedBasePath.endsWith(".rels")) {
    const relsMarker = "/_rels/";
    const relsMarkerIndex = normalizedBasePath.lastIndexOf(relsMarker);
    if (relsMarkerIndex >= 0) {
      const ownerPrefix = normalizedBasePath.slice(0, relsMarkerIndex);
      const relFileName = normalizedBasePath.slice(relsMarkerIndex + relsMarker.length);
      const ownerFileName = relFileName.endsWith(".rels")
        ? relFileName.slice(0, -".rels".length)
        : relFileName;
      baseDirectory = dirname(`${ownerPrefix}/${ownerFileName}`);
    }
  }

  const segments = [...baseDirectory.split("/").filter(Boolean), ...normalizedTarget.split("/").filter(Boolean)];
  const resolved: string[] = [];
  for (const segment of segments) {
    if (segment === ".") {
      continue;
    }
    if (segment === "..") {
      resolved.pop();
      continue;
    }
    resolved.push(segment);
  }

  return resolved.join("/");
}

export function readArchiveText(archive: Record<string, Uint8Array>, path: string | null | undefined) {
  if (!path) {
    return null;
  }

  const entry = archive[normalizeArchivePath(path)];
  return entry ? strFromU8(entry) : null;
}

export function parseXml(xml: string) {
  if (typeof DOMParser === "undefined") {
    return null;
  }

  try {
    return new DOMParser().parseFromString(xml, "application/xml");
  } catch {
    return null;
  }
}

export function serializeXml(document: XMLDocument) {
  return new XMLSerializer().serializeToString(document);
}

export function getLocalChildren(parent: ParentNode, localName: string) {
  return Array.from(parent.childNodes).filter(
    (node): node is Element => node.nodeType === Node.ELEMENT_NODE && (node as Element).localName === localName
  );
}

export function removeLocalChildren(parent: ParentNode, localName: string) {
  getLocalChildren(parent, localName).forEach((node) => node.parentNode?.removeChild(node));
}

export function getLocalDescendants(parent: ParentNode, localName: string) {
  return Array.from((parent as Element | Document).getElementsByTagName("*")).filter(
    (node) => node.localName === localName
  );
}

export function getFirstLocalChild(parent: ParentNode, localName: string) {
  return getLocalChildren(parent, localName)[0] ?? null;
}

export function getFirstLocalDescendant(parent: ParentNode, localName: string) {
  return getLocalDescendants(parent, localName)[0] ?? null;
}

export function ensureChild(parent: Element, localName: string, namespace = parent.namespaceURI ?? CHART_NS, prefix = "c") {
  const existing = getFirstLocalChild(parent, localName);
  if (existing) {
    return existing;
  }

  const document = parent.ownerDocument;
  const node = document.createElementNS(namespace, `${prefix}:${localName}`);
  parent.appendChild(node);
  return node;
}

export function setLeafValue(parent: Element, localName: string, value: string, namespace = parent.namespaceURI ?? CHART_NS, prefix = "c") {
  const node = ensureChild(parent, localName, namespace, prefix);
  node.textContent = value;
  return node;
}

export function setBooleanValue(parent: Element, localName: string, value: boolean) {
  const node = ensureChild(parent, localName);
  node.setAttribute("val", value ? "1" : "0");
  return node;
}

export function setNumericValue(parent: Element, localName: string, value: number) {
  const node = ensureChild(parent, localName);
  node.setAttribute("val", String(Math.round(value)));
  return node;
}

export function readChartNumericAttribute(parent: Element | null, localName: string) {
  const node = parent ? getFirstLocalChild(parent, localName) : null;
  const value = Number(node?.getAttribute("val") ?? Number.NaN);
  return Number.isFinite(value) ? value : undefined;
}

export function readChartBooleanAttribute(parent: Element | null, localName: string) {
  const node = parent ? getFirstLocalChild(parent, localName) : null;
  if (!node) {
    return undefined;
  }
  const rawValue = node.getAttribute("val");
  if (rawValue == null) {
    return true;
  }
  if (rawValue === "1" || rawValue === "true") {
    return true;
  }
  if (rawValue === "0" || rawValue === "false") {
    return false;
  }
  return undefined;
}
