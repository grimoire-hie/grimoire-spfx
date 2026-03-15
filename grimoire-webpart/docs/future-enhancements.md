# Future Enhancements

## True Intent Semantic Engine

### Why This Matters

The current search stack already has the right foundation:

- model-driven search planning
- semantic retrieval through Copilot Search and Copilot Retrieval
- lexical retrieval through classic SharePoint Search
- adaptive fan-out
- deterministic fusion and reranking

That is a strong V1, but it is not yet a true intent-semantic engine.

Today we generate:

- one semantic rewrite
- one SharePoint lexical query
- one correction candidate
- one translation fallback

This is useful, but still limited for ambiguous, multilingual, misspelled, or weakly phrased enterprise queries.

### Current Implementation Gap

The main gap is not basic planner/search coupling anymore. The main gaps are:

1. The planner still produces only a single interpretation path per branch.
2. Fallback execution is still driven mostly by result count rather than result quality.
3. Fusion reranks results, but it does not truly judge semantic relevance against the user intent.
4. There is no explicit ambiguity model or clarification mechanism.
5. Search quality still degrades noticeably when the fast planner times out or returns empty output.

### What A True Intent Semantic Engine Would Require

#### 1. Richer Intent Planning

The planner should evolve from a single-output rewrite service into a true intent planner.

Desired planner outputs:

- `queryLanguage`
- `semanticHypotheses[]`
- `sharePointLexicalHypotheses[]`
- `correctedQuery?`
- `translationFallbacks[]`
- `ambiguityScore`
- `clarificationQuestion?`
- `entities[]`
- `topics[]`
- per-hypothesis confidence

This would allow the system to search more than one plausible interpretation when the query is weak or ambiguous.

#### 2. Quality-Aware Orchestration

Search branching should not depend mainly on how many unique results were found.

The orchestrator should evaluate:

- source agreement
- semantic confidence
- lexical evidence quality
- outlier rate
- relevance of top results
- strength of same-language matches

This would allow the system to trigger fallback branches when the current results are broad but weak, not only when they are sparse.

#### 3. Semantic Reranking

This is the biggest missing capability.

The current system fuses ranked results from multiple APIs, but it does not explicitly ask:

> Is this result actually about what the user meant?

A stronger engine should add a second-stage semantic judge over the top fused candidates. That reranker should:

- compare result meaning to the user query or chosen hypothesis
- demote semantically adjacent but off-topic items
- reward direct topical matches
- preserve multilingual relevance when content is equivalent

#### 4. Ambiguity Handling

The system should distinguish between:

- clear intent
- partially expressed intent
- ambiguous intent
- low-confidence misspelled intent

When ambiguity is high and search quality is weak, the assistant should ask a short clarification question instead of pretending the best guess is reliable.

#### 5. Evaluation Infrastructure

This work should not be tuned ad hoc.

We need a repeatable evaluation set with:

- representative enterprise queries
- multilingual and misspelled examples
- expected strong matches
- expected false positives
- per-branch diagnostics
- before/after measurements

Without this, planner and reranker improvements will be difficult to validate objectively.

### Recommended Roadmap

#### Phase 1: Better Planner

- extend the search planner to emit multiple semantic and lexical hypotheses
- add ambiguity scoring
- add optional clarification output
- keep current raw query preservation

#### Phase 2: Better Branching

- replace count-driven fallback with quality-driven fallback
- trigger fallback branches based on weak source agreement or weak top-result quality
- keep SharePoint lexical-only and Copilot semantic-first

#### Phase 3: Semantic Reranker

- add a fast-model semantic reranker over the top fused results
- score results against the user query and active hypothesis
- use the reranker to demote topical neighbors and outliers

#### Phase 4: Evaluation Harness

- create a fixed query benchmark
- capture planner outputs and ranking explanations
- compare retrieval quality before and after each change

### Practical Conclusion

This is achievable in the current architecture. It does not require a new platform, but it does require more than:

- one planner prompt
- one lexical rewrite
- one fusion pass

The most important next capability is semantic reranking. That is the feature that turns the current search stack from a strong search orchestrator into a true intent-semantic engine.
