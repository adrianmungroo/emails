import win32com.client
import datetime
import csv
import json
import argparse
from pathlib import Path

OUTPUT_DIR = Path("output")

RANGES = ("today", "yesterday", "this-week", "last-week", "this-month", "last-month")

def resolve_range(name):
    today = datetime.date.today()
    if name == "today":
        return today, today + datetime.timedelta(days=1)
    if name == "yesterday":
        return today - datetime.timedelta(days=1), today
    if name == "this-week":        # week starts last Sunday
        start = today - datetime.timedelta(days=(today.weekday() + 1) % 7)
        return start, today + datetime.timedelta(days=1)
    if name == "last-week":
        this_sun = today - datetime.timedelta(days=(today.weekday() + 1) % 7)
        last_sun = this_sun - datetime.timedelta(weeks=1)
        return last_sun, this_sun
    if name == "this-month":
        return today.replace(day=1), today + datetime.timedelta(days=1)
    if name == "last-month":
        first = today.replace(day=1)
        prev  = first - datetime.timedelta(days=1)
        return prev.replace(day=1), first

def parse_args():
    p = argparse.ArgumentParser()
    g = p.add_mutually_exclusive_group()
    g.add_argument("--range", choices=RANGES, metavar=f"{{{','.join(RANGES)}}}")
    g.add_argument("--start")
    p.add_argument("--end")
    p.add_argument("--csv",  default="emails.csv")
    p.add_argument("--json", default="emails.json")
    return p.parse_args()

def fmt(dt): return dt.strftime("%m/%d/%Y %H:%M %p")

def main():
    args = parse_args()

    if args.range:
        s, e = resolve_range(args.range)
        start = datetime.datetime(s.year, s.month, s.day)
        end   = datetime.datetime(e.year, e.month, e.day)
    else:
        start = datetime.datetime.fromisoformat(args.start or "2026-03-10")
        end   = datetime.datetime.fromisoformat(args.end   or "2026-03-18")

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    messages = outlook.GetDefaultFolder(6).Items
    messages.Sort("[ReceivedTime]", True)

    filtered = messages.Restrict(
        f"[ReceivedTime] >= '{fmt(start)}' AND [ReceivedTime] < '{fmt(end)}'"
    )

    OUTPUT_DIR.mkdir(exist_ok=True)
    results = []
    for msg in filtered:
        if getattr(msg, "Class", None) != 43:  # skip non-mail (calendar, etc.)
            continue
        results.append({
            "subject":       msg.Subject,
            "sender":        msg.SenderName,
            "sender_email":  msg.SenderEmailAddress,
            "received_time": str(msg.ReceivedTime),
            "body":          msg.Body.strip(),
        })

    csv_path  = OUTPUT_DIR / args.csv
    json_path = OUTPUT_DIR / args.json

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=results[0].keys())
        writer.writeheader()
        writer.writerows(results)

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    print(f"Saved {len(results)} emails → {csv_path}, {json_path}")

if __name__ == "__main__":
    main()
