/**
 * Model indexing: builds the in-memory `ModelIndex` (class buckets, level
 * buckets, item→level map, item names, spatial browser tree) from a loaded
 * fragments model. Extracted from viewer.ts — it reads only the model (via the
 * FragmentsModelLike boundary) and the pure property-engine label/storey
 * helpers, and returns a fresh ModelIndex; the orchestrator stores it. No DOM,
 * no app state.
 */
import type { SpatialTreeItem } from '@thatopen/fragments';
import { extractStoreyNameFromItemData, getModelTreeItemLabel } from './property-engine';
import type { FragmentsModelLike } from './fragments-model';
import type { BrowserTreeNode, ModelIndex } from './viewer-types';

// Chunk size for the getItemsData round-trips (P7: batched, not per-item).
const CHUNK_SIZE = 360;

function collectSpatialTreeIds(node: SpatialTreeItem, target: Set<number>): void {
  if (node.localId !== null) target.add(node.localId);
  for (const child of node.children ?? []) collectSpatialTreeIds(child, target);
}

function buildSpatialBrowserTree(
  node: SpatialTreeItem,
  itemNames: Map<number, string>,
  geometryIds: Set<number>,
): BrowserTreeNode {
  const localId = node.localId;
  const category = node.category ?? 'Group';
  const children = (node.children ?? []).map((child) => buildSpatialBrowserTree(child, itemNames, geometryIds));

  let geometryCount = localId !== null && geometryIds.has(localId) ? 1 : 0;
  for (const child of children) geometryCount += child.geometryCount;

  const label = localId !== null
    ? (itemNames.get(localId) ?? `${category} ${localId}`)
    : (category || 'Structure');

  return {
    category,
    localId,
    label,
    geometryCount,
    children,
  };
}

/**
 * Builds the full model index for a loaded fragments model. Behaviour matches
 * the pre-extraction inline `indexModel` exactly (same chunking, same fallback
 * storey walk, same tree build).
 */
export async function buildModelIndex(modelId: string, model: FragmentsModelLike): Promise<ModelIndex> {
  const itemIds = await model.getItemsIdsWithGeometry();
  const idsSet = new Set(itemIds);

  const classes = new Map<string, Set<number>>();
  const itemClassById = new Map<number, string>();
  const categories = await model.getItemsWithGeometryCategories();
  for (let i = 0; i < categories.length; i += 1) {
    const category = categories[i] ?? 'Unknown';
    const id = itemIds[i];
    if (typeof id !== 'number') continue;
    if (!classes.has(category)) classes.set(category, new Set<number>());
    classes.get(category)?.add(id);
    itemClassById.set(id, category);
  }

  const itemNames = new Map<number, string>();
  const itemToLevel = new Map<number, string>();
  const levels = new Map<string, Set<number>>();
  const spatial = await model.getSpatialStructure();

  // Read element names + level assignment from ContainedInStructure relation.
  // P7 (W5.3): the chunk round-trips to the fragments worker are independent, so
  // fire them in PARALLEL (Promise.all) instead of awaiting one at a time —
  // hundreds of sequential round-trips were the dominant indexing cost. The
  // per-chunk results are then applied in chunk order so the maps are identical
  // to the sequential build (deterministic).
  const chunks: number[][] = [];
  for (let start = 0; start < itemIds.length; start += CHUNK_SIZE) {
    chunks.push(itemIds.slice(start, start + CHUNK_SIZE));
  }
  const chunkResults = await Promise.all(
    chunks.map((chunk) =>
      model.getItemsData(chunk, {
        attributesDefault: true,
        relations: {
          ContainedInStructure: { attributes: true, relations: true },
        },
        relationsDefault: { attributes: false, relations: false },
      }),
    ),
  );
  for (let c = 0; c < chunks.length; c += 1) {
    const chunk = chunks[c];
    const itemsData = chunkResults[c];
    for (let i = 0; i < chunk.length; i += 1) {
      const localId = chunk[i];
      const data = (itemsData[i] || {}) as Record<string, unknown>;
      const category = itemClassById.get(localId) ?? 'Element';
      itemNames.set(localId, getModelTreeItemLabel(data, localId, category));
      const levelName = extractStoreyNameFromItemData(data);
      if (!levelName) continue;
      itemToLevel.set(localId, levelName);
      if (!levels.has(levelName)) levels.set(levelName, new Set<number>());
      levels.get(levelName)?.add(localId);
    }
  }

  // Spatial fallback: ensure storey names are loaded and assign ungrouped items.
  const storeyIds = new Set<number>();
  const collectStoreys = (node: SpatialTreeItem): void => {
    const category = (node.category ?? '').toUpperCase();
    if (category.includes('IFCBUILDINGSTOREY') && node.localId !== null) storeyIds.add(node.localId);
    for (const child of node.children ?? []) collectStoreys(child);
  };
  collectStoreys(spatial);

  const unknownStoreyIds = [...storeyIds].filter((id) => !itemNames.has(id));
  for (let start = 0; start < unknownStoreyIds.length; start += CHUNK_SIZE) {
    const chunk = unknownStoreyIds.slice(start, start + CHUNK_SIZE);
    const rows = await model.getItemsData(chunk, {
      attributesDefault: true,
      relationsDefault: { attributes: false, relations: false },
    });
    for (let i = 0; i < chunk.length; i += 1) {
      const localId = chunk[i];
      const data = (rows[i] || {}) as Record<string, unknown>;
      itemNames.set(localId, getModelTreeItemLabel(data, localId, 'Storey'));
    }
  }

  const walkSpatial = (node: SpatialTreeItem, activeStorey: string | null): void => {
    const category = (node.category ?? '').toUpperCase();
    let nextStorey = activeStorey;
    if (category.includes('IFCBUILDINGSTOREY') && node.localId !== null) {
      nextStorey = itemNames.get(node.localId) ?? `Storey ${node.localId}`;
    }
    if (node.localId !== null && nextStorey && idsSet.has(node.localId) && !itemToLevel.has(node.localId)) {
      itemToLevel.set(node.localId, nextStorey);
      if (!levels.has(nextStorey)) levels.set(nextStorey, new Set<number>());
      levels.get(nextStorey)?.add(node.localId);
    }
    for (const child of node.children ?? []) walkSpatial(child, nextStorey);
  };
  walkSpatial(spatial, null);

  // Load names for non-geometry nodes used by the browser spatial tree.
  const spatialIds = new Set<number>();
  collectSpatialTreeIds(spatial, spatialIds);
  const missingSpatialIds = [...spatialIds].filter((id) => !itemNames.has(id));
  for (let start = 0; start < missingSpatialIds.length; start += CHUNK_SIZE) {
    const chunk = missingSpatialIds.slice(start, start + CHUNK_SIZE);
    const rows = await model.getItemsData(chunk, {
      attributesDefault: true,
      relationsDefault: { attributes: false, relations: false },
    });
    for (let i = 0; i < chunk.length; i += 1) {
      const localId = chunk[i];
      const data = (rows[i] || {}) as Record<string, unknown>;
      itemNames.set(localId, getModelTreeItemLabel(data, localId, 'Item'));
    }
  }

  const spatialRoot = buildSpatialBrowserTree(spatial, itemNames, idsSet);

  return {
    modelId,
    allIds: new Set(itemIds),
    classes,
    levels,
    itemToLevel,
    itemNames,
    spatialRoot,
  };
}
