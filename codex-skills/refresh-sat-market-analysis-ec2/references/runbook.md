# EC2 Market Analysis Refresh Runbook

## 1. Initial Inspection

```sh
ssh ec2-devdesk-nakai-via-vpn
cd /home/ryo_saito_s/sat
hostname
git status --short
git branch --show-current
git rev-parse --short HEAD
git fetch --all --prune
git log --oneline -1 HEAD
git log --oneline -1 origin/feature/market_analysis_agent
git log --oneline -1 origin/feat/market-analysis-fe
```

Explain to the user that this EC2 setup is intentionally branch-hybrid: PA branch as repo base, FE directory from FE branch.

## 2. Backup Before Update

Use a timestamped directory. Capture both tracked and untracked work.

```sh
cd /home/ryo_saito_s/sat
backup=/home/ryo_saito_s/sat-update-backup-$(date +%Y%m%d_%H%M%S)
mkdir -p "$backup"
git status --short > "$backup/status.txt"
git diff > "$backup/unstaged.patch"
git diff --staged > "$backup/staged.patch"
git ls-files --others --exclude-standard > "$backup/untracked-files.txt"
tar -czf "$backup/untracked-files.tar.gz" -T "$backup/untracked-files.txt" 2>/dev/null || true
git rev-parse HEAD > "$backup/head-before.txt"
git rev-parse origin/feature/market_analysis_agent > "$backup/pa-origin-before.txt"
git rev-parse origin/feat/market-analysis-fe > "$backup/fe-origin-before.txt"
echo "$backup"
```

## 3. Update Source

PA base update:

```sh
cd /home/ryo_saito_s/sat
git pull --ff-only
```

If fast-forward is not possible, stop and inspect. Do not reset.

FE branch content update:

```sh
cd /home/ryo_saito_s/sat
git restore --source origin/feat/market-analysis-fe --staged --worktree -- src/sat/frontend
```

Then check whether EC2-only fixes are still required. Pay special attention to fixes that were present in the pre-update backup; restoring FE from the FE branch may remove them.

```sh
git cat-file -t 3080ff27d 2>/dev/null || true
git diff --name-only origin/feat/market-analysis-fe -- \
  src/sat/frontend/server/utils/paAuthHeaders.ts \
  src/sat/frontend/server/api/market-sizing \
  src/sat/frontend/server/api/forced-ideation \
  src/sat/frontend/app/plugins/02.authRedirect.ts \
  src/sat/frontend/Dockerfile

# If a backup exists, check whether the previous EC2 state already had HTTP-cookie/auth/build fixes.
latest_backup=$(ls -dt /home/ryo_saito_s/sat-update-backup-* 2>/dev/null | head -n 1 || true)
if [ -n "$latest_backup" ]; then
  grep -n "02.authRedirect\|ui=agent-cockpit\|SameSite=Lax; Secure\|paAuthHeaders\|NODE_OPTIONS" "$latest_backup"/*.patch 2>/dev/null || true
fi
```

Reapply FE auth fix from the existing commit only if the FE branch does not already contain equivalent code:

```sh
git restore --source 3080ff27d --staged --worktree -- \
  src/sat/frontend/server/utils/paAuthHeaders.ts \
  src/sat/frontend/server/api/forced-ideation/check.post.ts \
  src/sat/frontend/server/api/forced-ideation/index.post.ts \
  src/sat/frontend/server/api/forced-ideation/stream.get.ts \
  src/sat/frontend/server/api/market-sizing/check-topic.post.ts \
  src/sat/frontend/server/api/market-sizing/generate.post.ts \
  src/sat/frontend/server/api/market-sizing/sessions/index.get.ts \
  src/sat/frontend/server/api/market-sizing/sessions/[id]/index.get.ts \
  src/sat/frontend/server/api/market-sizing/sessions/[id]/judge.post.ts \
  src/sat/frontend/server/api/market-sizing/sessions/[id]/patch-apply.post.ts \
  src/sat/frontend/server/api/market-sizing/sessions/[id]/patch.post.ts \
  src/sat/frontend/server/api/market-sizing/sessions/[id]/stream.get.ts
```

## 4. Compose Fixes If Still Needed

### FE routing cookie on EC2 HTTP

This incident can happen when the FE directory is restored from `origin/feat/market-analysis-fe` and an EC2-only HTTP workaround is overwritten. Treat this as a source-update side effect, not as a Docker or VPN problem first.

Check `src/sat/frontend/app/plugins/02.authRedirect.ts`. The EC2 demo is served via `http://10.0.10.184`, so `ui=agent-cockpit` must not always be written with `Secure`.

Problematic branch form:

```ts
document.cookie =
  'ui=agent-cockpit; path=/; max-age=86400; SameSite=Lax; Secure';
window.location.reload();
```

EC2-safe form:

```ts
const secureCookie = window.location.protocol === 'https:' ? '; Secure' : '';
document.cookie =
  'ui=agent-cockpit; path=/; max-age=86400; SameSite=Lax' + secureCookie;
window.location.reload();
```

After changing this, rebuild FE:

```sh
cd /home/ryo_saito_s/sat/src/sat
docker compose up -d --build fe
```

Verify the built output and watch for repeated reload-like access:

```sh
docker exec sat-fe-1 sh -lc 'grep -R "ui=agent-cockpit; path=/; max-age=86400; SameSite=Lax" -n /app/.output/public /app/.output/server 2>/dev/null | head'
docker compose logs --since=30m local-proxy | grep -v Wget | tail -n 120
```

Completion is strongest when the affected user retries and the `local-proxy` logs no longer show repeated `GET /`, `/agents`, `/api/auth/me`, or `/api/agents` loops at that time.

### FE build memory

If Nuxt build hits heap OOM, add this in `src/sat/frontend/Dockerfile` build stage after the Corepack env lines:

```dockerfile
ENV NODE_OPTIONS=--max-old-space-size=4096
```

### PA S3 settings for forced ideation

Check PA env in compose and in the running container:

```sh
cd /home/ryo_saito_s/sat
awk '/^  pa:/{flag=1} flag{print} /^  agw:/{if(flag) exit}' src/sat/docker-compose.yml | sed -n '1,100p'
docker exec sat-pa-1 sh -lc 'env | sort | grep -E "SAT_S3_CUSTOMER_BUCKET|AWS_ENDPOINT_URL_S3|AWS_DEFAULT_REGION|AWS_ACCESS_KEY_ID|AWS_SECRET_ACCESS_KEY" || true'
```

If missing, add under `pa.environment`:

```yaml
      # S3-compatible local storage for forced ideation job artifacts
      SAT_S3_CUSTOMER_BUCKET: ${SAT_BUCKET_NAME:-sat-local-bucket}
      AWS_ENDPOINT_URL_S3: http://seaweedfs:8333
      AWS_DEFAULT_REGION: ap-northeast-1
      AWS_ACCESS_KEY_ID: admin
      AWS_SECRET_ACCESS_KEY: admin
```

Then recreate PA:

```sh
cd /home/ryo_saito_s/sat/src/sat
docker compose up -d pa
```

Verify:

```sh
docker exec sat-pa-1 sh -lc 'python - <<PY
from professional_agents.forced_ideation.services.job_storage import JobStorage
s=JobStorage()
print(s.status_key("tenant", "user", "job"))
PY'

docker exec sat-pa-1 sh -lc 'python - <<PY
import boto3, os
s3 = boto3.client("s3")
bucket = os.environ["SAT_S3_CUSTOMER_BUCKET"]
s3.head_bucket(Bucket=bucket)
print(f"bucket ok: {bucket}")
PY'
```

## 5. Build / Restart

Full refresh:

```sh
cd /home/ryo_saito_s/sat/src/sat
docker compose up -d --build pa fe
```

If only FE changed:

```sh
cd /home/ryo_saito_s/sat/src/sat
docker compose up -d --build fe
```

If only PA environment changed:

```sh
cd /home/ryo_saito_s/sat/src/sat
docker compose up -d pa
```

## 6. Verification Commands

Source state:

```sh
cd /home/ryo_saito_s/sat
echo "PA HEAD: $(git rev-parse --short HEAD) $(git log -1 --format=%s)"
echo "PA origin: $(git rev-parse --short origin/feature/market_analysis_agent) $(git log -1 --format=%s origin/feature/market_analysis_agent)"
echo "FE origin: $(git rev-parse --short origin/feat/market-analysis-fe) $(git log -1 --format=%s origin/feat/market-analysis-fe)"
echo -n "PA market-related diff files: "
git diff --name-only origin/feature/market_analysis_agent -- src/sat/backend/professional_agents/market_analysis src/sat/backend/professional_agents/market_sizing src/sat/backend/professional_agents/forced_ideation | wc -l
echo -n "FE diff files vs FE branch: "
git diff --name-only origin/feat/market-analysis-fe -- src/sat/frontend | wc -l
```

Remember: FE diff may be nonzero because of intentional EC2-only fixes.

Image/container match:

```sh
cd /home/ryo_saito_s/sat/src/sat
docker image inspect sat-fe sat-pa --format '{{.RepoTags}} created={{.Created}} id={{.Id}}'
docker inspect sat-fe-1 sat-pa-1 --format '{{.Name}} image={{.Image}} created={{.Created}} started={{.State.StartedAt}}'
docker compose ps fe pa aa dca agw
```

HTTP checks:

```sh
curl -s -o /dev/null -w 'fe %{http_code}\n' http://localhost:3000/agents
curl -s -o /dev/null -w 'pa %{http_code}\n' http://localhost:8004/health
curl -s -o /dev/null -w 'aa %{http_code}\n' http://localhost:8006/health
```

## 7. Known Symptoms

- `401 Unauthorized`, `Provide Authorization header or X-API-Key`:
  - FE server route is not forwarding `Authorization: Bearer <token>` to PA.
  - Check `paAuthHeaders` and the market-sizing / forced-ideation routes.
- `Failed to start forced ideation` with PA log `SAT_S3_CUSTOMER_BUCKET is not set`:
  - Add/reapply PA S3 environment variables and recreate PA.
- Nuxt build OOM:
  - Add/reapply `NODE_OPTIONS=--max-old-space-size=4096` and rebuild FE.
- External user sees repeated reloads or the agent list never appears, while your browser works:
  - Do not conclude from your browser alone. Check `local-proxy` logs for that user's retry time.
  - If logs show repeated `GET /`, `/agents`, `/api/auth/me`, `/api/agents`, or many `499` cancels, check the FE routing cookie fix in `02.authRedirect.ts`.
  - Rebuild FE after applying the HTTP-safe cookie change. Ask the affected user to retry; that user's retry is the real completion signal.
- `docker compose` warnings for Databricks vars:
  - Seen during this EC2 setup; not the cause of the market analysis agent failures unless SQL RAG is being tested.

## 8. Slack Report Template

```text
@中井悠/Haruka Nakai
EC2上の市場分析環境の最新化と動作確認が完了しました。

実施内容:
- PA側を feature/market_analysis_agent の最新に更新
- FE側を feat/market-analysis-fe の最新に更新
- Dockerイメージを再ビルドし、コンテナを再起動
- 市場分析エージェント、強制発想法、市場規模算出エージェントの起動・基本動作を確認

補足:
- FEからPAへの認証ヘッダー不足があったため、EC2上で追加対応しました
- 強制発想法の保存先設定がPAコンテナに不足していたため、ローカルS3設定を追加しました
- 現在、主要コンテナは healthy です

ご確認よろしくお願いします。
```

For reload-loop incidents, use a more cautious report:

```text
EC2上の市場分析環境について、リロードが繰り返される可能性があるFE側のCookie設定を修正し、FEコンテナを再ビルド・再起動しました。

現在、主要コンテナは正常起動しており、/agents も正常応答することを確認しています。
お手数ですが、再度 http://10.0.10.184/agents をお試しください。

もしまだリロードが続く場合は、発生時刻を教えてください。こちらでその時刻のログを確認します。
```
