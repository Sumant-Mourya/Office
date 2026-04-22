import json
import os
from datetime import datetime
from google import genai
import config

TRACK_FILE = "batch_tracking.json"
client = genai.Client(api_key=config.GEMINI_API_KEY)


def load_tracking():
    if not os.path.exists(TRACK_FILE):
        return []
    with open(TRACK_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_tracking(data):
    with open(TRACK_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)


def extract_tokens(job):
    """Try to pull token usage from the batch job object."""
    try:
        dump = job.model_dump() if hasattr(job, "model_dump") else {}
    except Exception:
        dump = {}

    # Gemini may expose usage_metadata at top level or inside nested fields
    usage = dump.get("usage_metadata") or dump.get("usageMetadata") or {}
    if not usage and isinstance(dump, dict):
        # walk one level deep
        for v in dump.values():
            if isinstance(v, dict):
                usage = v.get("usage_metadata") or v.get("usageMetadata") or {}
                if usage:
                    break

    prompt_tokens = usage.get("prompt_token_count") or usage.get("promptTokenCount") or 0
    output_tokens = usage.get("candidates_token_count") or usage.get("candidatesTokenCount") or 0
    total_tokens = usage.get("total_token_count") or usage.get("totalTokenCount") or (prompt_tokens + output_tokens)

    return prompt_tokens, output_tokens, total_tokens


def monitor():
    tracking = load_tracking()

    if not tracking:
        print("\n" + "=" * 60)
        print("  No batches found in tracking.")
        print("  Run batch_submit.py first to create batches.")
        print("=" * 60 + "\n")
        return

    # Counters
    succeeded = 0
    running = 0
    failed = 0
    pending = 0
    cancelled = 0
    already_downloaded = 0
    total_prompt_tokens = 0
    total_output_tokens = 0
    total_tokens_all = 0

    print("\n" + "=" * 70)
    print(f"  {'Batch':<8} {'Status':<25} {'Rows':<7} {'Tokens':>12}  {'Job ID'}")
    print("-" * 70)

    for entry in tracking:
        batch_num = entry.get("batch", entry.get("batch_number", "?"))
        rows = entry.get("rows", "?")
        job_id = entry.get("job_id")

        if not job_id:
            entry["status"] = "NO_JOB"
            print(f"  {batch_num:<8} {'NO JOB':<25} {str(rows):<7} {'—':>12}  —")
            pending += 1
            continue

        if entry.get("downloaded"):
            status = entry.get("status", "JOB_STATE_SUCCEEDED")
            pt = entry.get("prompt_tokens", 0)
            ot = entry.get("output_tokens", 0)
            tt = entry.get("total_tokens", pt + ot)
            token_str = str(tt) if tt else "—"
            total_prompt_tokens += pt
            total_output_tokens += ot
            total_tokens_all += tt
            already_downloaded += 1
            succeeded += 1
            print(f"  {batch_num:<8} {'DOWNLOADED':<25} {str(rows):<7} {token_str:>12}  {job_id}")
            continue

        # Query the API
        try:
            job = client.batches.get(name=job_id)
            state = job.state.name
        except Exception as e:
            state = f"ERROR: {e}"
            entry["status"] = state
            print(f"  {batch_num:<8} {'QUERY FAILED':<25} {str(rows):<7} {'—':>12}  {job_id}")
            failed += 1
            continue

        entry["status"] = state

        # Token usage
        pt, ot, tt = extract_tokens(job)
        if tt:
            entry["prompt_tokens"] = pt
            entry["output_tokens"] = ot
            entry["total_tokens"] = tt
            total_prompt_tokens += pt
            total_output_tokens += ot
            total_tokens_all += tt
        token_str = str(tt) if tt else "—"

        # Handle result file for succeeded jobs
        if state == "JOB_STATE_SUCCEEDED":
            entry["completed_at"] = entry.get("completed_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            # Try to get result file
            try:
                if hasattr(job, "dest") and job.dest:
                    entry["result_file"] = getattr(job.dest, "file_name", None) or getattr(job.dest, "name", None)
            except Exception:
                pass
            succeeded += 1

        elif state == "JOB_STATE_FAILED":
            entry["failed"] = True
            entry["completed_at"] = entry.get("completed_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            failed += 1

        elif state in ("JOB_STATE_CANCELLED", "JOB_STATE_CANCELLING"):
            cancelled += 1

        elif state in ("JOB_STATE_RUNNING", "JOB_STATE_PENDING"):
            running += 1

        else:
            pending += 1

        # Friendly label
        label = state.replace("JOB_STATE_", "")
        print(f"  {batch_num:<8} {label:<25} {str(rows):<7} {token_str:>12}  {job_id}")

    print("-" * 70)

    save_tracking(tracking)

    # ── Summary ──────────────────────────────────────────────────
    total_batches = len(tracking)
    total_rows = sum(e.get("rows", 0) for e in tracking)

    print(f"\n{'=' * 60}")
    print(f"  Monitor Summary")
    print(f"{'=' * 60}")
    print(f"  Total batches      : {total_batches}")
    print(f"  Total rows         : {total_rows}")
    print(f"  ---")
    print(f"  Succeeded          : {succeeded}")
    print(f"  Running / Pending  : {running}")
    print(f"  Failed             : {failed}")
    print(f"  Cancelled          : {cancelled}")
    print(f"  Already downloaded : {already_downloaded}")

    if total_tokens_all:
        print(f"  ---")
        print(f"  Prompt tokens      : {total_prompt_tokens:,}")
        print(f"  Output tokens      : {total_output_tokens:,}")
        print(f"  Total tokens       : {total_tokens_all:,}")

    if running > 0:
        print(f"\n  {running} batch(es) still running — re-run monitor later.")
    elif failed > 0 and succeeded == 0:
        print(f"\n  All batches failed!")
    elif running == 0 and failed == 0 and succeeded > 0:
        ready = succeeded - already_downloaded
        if ready > 0:
            print(f"\n  All done! {ready} batch(es) ready to download.")
            print(f"  Run batch_download.py to download results.")
        else:
            print(f"\n  All batches downloaded. Nothing to do.")

    print(f"{'=' * 60}\n")


if __name__ == "__main__":
    monitor()
