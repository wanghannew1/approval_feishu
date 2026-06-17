# Bugfix Learnings

## Bug 1: Attachment Download 404
- Root cause: attachmentV2 value returns full temporary URL (12h TTL), not file_token
- Fix: Auto-detect format — starts with http → use directly, else → drive API
- Result: 171.5 KB Excel downloaded successfully, no 404

## Bug 2: Search API Validation Error
- First attempt: Added user_id/offset/limit/sort_asc to /search endpoint → 60001/99992402 errors
- Actual fix: Wrong endpoint. Changed to /instances/query with page_size/page_token
- Note: /query endpoint needs approval:approval.list:readonly scope
- Result: Query succeeds, returns 0 APPROVED (only instance is PENDING)

## Lessons
- Never assume API behavior based on docs alone — verify with real calls
- When adding "required" params doesn't fix it, check if endpoint is wrong
