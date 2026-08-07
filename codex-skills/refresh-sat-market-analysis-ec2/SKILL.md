---
name: refresh-sat-market-analysis-ec2
description: Use when updating or verifying the EC2 market analysis demo environment reached via `ssh ec2-devdesk-nakai-via-vpn`, especially when refreshing PA `feature/market_analysis_agent` and FE `feat/market-analysis-fe`, rebuilding Docker containers, troubleshooting market analysis / market sizing / forced ideation, or preparing a concise Slack status for Nakai-san.
metadata:
  short-description: Refresh and verify the SAT market analysis EC2 demo
---

# Refresh SAT Market Analysis EC2

Use this skill for the repeat task of updating the EC2 demo environment for the SAT market analysis agents.

## Core Context

- SSH host: `ec2-devdesk-nakai-via-vpn`
- EC2 repo: `/home/ryo_saito_s/sat`
- Compose dir: `/home/ryo_saito_s/sat/src/sat`
- PA target branch: `feature/market_analysis_agent`
- FE target branch: `feat/market-analysis-fe`
- Typical URL: `http://10.0.10.184/agents`

The environment is a hybrid: the repo's current branch is usually the PA branch, while `src/sat/frontend` is intentionally restored from the FE branch. Do not assume a single Git branch fully describes the working tree.

## Operating Rules

- Explain in plain Japanese. Avoid unclear terms like "checkout" or "local adjustment"; prefer "FE側の最新内容を反映" and "EC2環境用の追加対応".
- Be careful, but do not ask for confirmation at every small step. Ask only before destructive actions, risky rollback, or changing remote state in a non-obvious way.
- Always back up current EC2 diffs before updating.
- Treat existing EC2 diffs as potentially intentional environment fixes. Before restoring FE from the FE branch, check whether those diffs contain EC2 HTTP, auth, S3, or build workarounds that must be preserved.
- Never discard unrelated EC2 changes without a backup and an explicit reason.
- For verification, think in three layers: source files -> Docker image -> running container. Docker images and containers do not remember branch names.

## Standard Workflow

1. Connect and inspect:
   - SSH to the host.
   - Confirm repo path, current branch, current commit, remotes, and `git status`.
   - Fetch remotes.
2. Back up current EC2 changes:
   - Save `git status`, staged/unstaged patches, untracked file list, and untracked archive under `/home/ryo_saito_s/sat-update-backup-YYYYMMDD_HHMMSS`.
3. Update source:
   - Fast-forward PA branch to `origin/feature/market_analysis_agent` when possible.
   - Restore `src/sat/frontend` from `origin/feat/market-analysis-fe`.
   - Preserve or reapply EC2-only fixes listed below if still needed.
4. Build and restart:
   - From `/home/ryo_saito_s/sat/src/sat`, rebuild/recreate `pa` and `fe` as needed.
   - If only compose environment changed for PA, recreating PA may be enough.
5. Verify:
   - Confirm source diffs against both target branches.
   - Confirm image IDs and running container image IDs match.
   - Confirm `fe`, `pa`, `aa`, `dca`, `agw` are healthy.
   - Confirm HTTP endpoints and the browser-visible smoke tests.
6. Report:
   - Summarize target branch commits, EC2-only additions, container health, and manual UI checks.

For exact commands and known failures, read `references/runbook.md`.

## Known EC2-Only Fixes From 2026-07-01

These may become unnecessary once upstream branches include them. Check before reapplying.

- FE -> PA auth forwarding:
  - Symptom: market sizing or forced ideation returns 401 with `Provide Authorization header or X-API-Key`.
  - Existing fix commit: `3080ff27d fix: [FE] SD-249: BFF から PA 呼び出しに Authorization を一律付与 (#4224)`.
  - Relevant files include `src/sat/frontend/server/utils/paAuthHeaders.ts`, `server/api/market-sizing/**`, and `server/api/forced-ideation/**`.
- FE build memory:
  - Symptom: Nuxt build fails with JavaScript heap out of memory.
  - EC2 workaround: set `ENV NODE_OPTIONS=--max-old-space-size=4096` in `src/sat/frontend/Dockerfile` build stage.
- Forced ideation S3 settings for PA:
  - Symptom: `Failed to start forced ideation`; PA logs show `RuntimeError: SAT_S3_CUSTOMER_BUCKET is not set`.
  - EC2 workaround: add PA env vars in `src/sat/docker-compose.yml`: `SAT_S3_CUSTOMER_BUCKET`, `AWS_ENDPOINT_URL_S3`, `AWS_DEFAULT_REGION`, `AWS_ACCESS_KEY_ID`, `AWS_SECRET_ACCESS_KEY`.
- FE Agent Cockpit routing cookie on EC2 HTTP:
  - Symptom: external users see repeated reloads or cannot reach the agent list, while `/agents` may work for users whose browser state already avoids the path. `local-proxy` logs may show repeated `GET /`, `/agents`, `/api/auth/me`, `/api/agents`, or `499` cancels.
  - Cause: restoring FE from `origin/feat/market-analysis-fe` can overwrite an EC2-only fix in `src/sat/frontend/app/plugins/02.authRedirect.ts`. The branch version sets `ui=agent-cockpit` with `Secure` unconditionally; the EC2 URL is HTTP, so this can break cookie persistence or trigger reload behavior for some browsers.
  - EC2 workaround: set the `Secure` attribute only when `window.location.protocol === 'https:'`, then rebuild FE.

## UI Smoke Tests

- FE latest visual marker:
  - Market analysis company placeholder is `例：ストックマーク株式会社`.
  - Important material upload is a paperclip-style `ファイルを添付` button.
- Functional checks:
  - Market analysis: input flow starts and proceeds.
  - Market sizing: opens, topic confirmation/generation starts, session appears.
  - Forced ideation: input check and `作成` starts without 401/500.
