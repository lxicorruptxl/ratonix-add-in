export type Tone = "balanced" | "softer" | "firmer";

export type DraftKind = "rewrite" | "reply";

export type SuggestionSource = "api" | "local";

export interface MessageSnapshot {
  bodyText: string;
  subject: string;
  from?: string;
  to: string[];
  conversationId?: string;
  mode: "compose" | "read" | "unknown";
}

export interface RelationshipMemory {
  address: string;
  displayName?: string;
  notes: string[];
  interactionCount: number;
  updatedAt: string;
}

export interface StyleMemory {
  examples: string[];
  interactions: Array<{
    tone: Tone;
    mode: DraftKind;
    createdAt: string;
    subject: string;
    sourcePreview: string;
    suggestionPreview: string;
  }>;
  keyRelationships: string[];
  relationships: Record<string, RelationshipMemory>;
}

export interface RatonixSuggestion {
  text: string;
  rationale: string;
  source: SuggestionSource;
}

export interface RatonixRequest {
  message: MessageSnapshot;
  tone: Tone;
  memory: StyleMemory;
}
