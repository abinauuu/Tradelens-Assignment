import argparse, csv, json, os, warnings
from collections import OrderedDict
from datetime import datetime, timedelta
import openpyxl

IST = timedelta(hours=5, minutes=30)
COLS = ["trade_id", "source", "symbol", "asset_class", "close_date", "close_time",
        "quantity", "pnl_native", "currency", "pnl_inr"]

def asset_class(sym):
    for p, c in (("XAU", "metals"), ("XAG", "metals"), ("EUR", "forex"), ("GBP", "forex"),
                 ("USD", "forex"), ("BTC", "crypto"), ("ETH", "crypto"),
                 ("BANKNIFTY", "index_options"), ("FINNIFTY", "index_options"),
                 ("NIFTY", "index_options")):
        if sym.startswith(p):
            return c
    return "other"

def clean_symbol(s):
    return str(s).strip().upper().split(".")[0]


def num(v):
    if v is None or (isinstance(v, str) and not v.strip()):
        return 0.0
    if isinstance(v, str):
        v = v.replace(",", "").strip()
    return float(v)


def read_mt5(path, usd_inr):
    out = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            r = {(k or "").strip(): v for k, v in r.items()}
            if (r.get("Type") or "").strip().lower() not in ("buy", "sell"):
                continue
            t = datetime.strptime(r["Close Time"].strip(), "%Y.%m.%d %H:%M:%S") + IST
            pnl = round(num(r["Profit"]) + num(r["Commission"]) + num(r["Swap"]), 2) + 0.0
            sym = clean_symbol(r["Symbol"])
            out.append({"trade_id": "mt5-" + r["Ticket"].strip(), "source": "mt5", "symbol": sym,
                        "asset_class": asset_class(sym), "close_date": t.strftime("%Y-%m-%d"),
                        "close_time": t.strftime("%H:%M:%S"), "quantity": num(r["Volume"]),
                        "pnl_native": pnl, "currency": "USD", "pnl_inr": round(pnl * usd_inr, 2) + 0.0})
    return out


def read_broker(path):
    out = []
    rows = []
    for ws in openpyxl.load_workbook(path, data_only=True).worksheets:
        rows.extend(ws.iter_rows(values_only=True))

    h = next(
        (i for i, r in enumerate(rows)
         if r and any(str(c).strip().lower() == "order id"
                      for c in r if c is not None)),
        None
    )
    if h is None:
        raise ValueError("Could not find 'Order ID' header in " + path)

    head = [str(c).strip() if c is not None else "" for c in rows[h]]
    for r in rows[h + 1:]:
        d = dict(zip(head, r))
        oid = d.get("Order ID")
        if oid is None or str(oid).strip() == "" or str(oid).strip().lower() == "total":
            continue
        if d.get("Realized P&L") in (None, ""):
            continue
        dt = d["Trade Date"]
        if isinstance(dt, datetime):
            pass
        elif isinstance(dt, (int, float)):
            dt = openpyxl.utils.datetime.from_excel(dt)
        else:
            dt = datetime.strptime(str(dt).strip(), "%d-%m-%Y")
            pnl = round(num(d["Realized P&L"]), 2) + 0.0
            sym = clean_symbol(d["Symbol"])
            q = num(d["Quantity"])
            out.append({"trade_id": "in-%s" % str(oid).strip().split(".")[0], "source": "broker_in",
                        "symbol": sym, "asset_class": asset_class(sym),
                        "close_date": dt.strftime("%Y-%m-%d"), "close_time": "",
                        "quantity": int(q) if q == int(q) else q, "pnl_native": pnl,
                        "currency": "INR", "pnl_inr": pnl})
            return out


def load_trades(data_dir, usd_inr):
    seen = OrderedDict()
    for name in sorted(os.listdir(data_dir)):
        p = os.path.join(data_dir, name)
        low = name.lower()
        if low.endswith(".csv"):
            rows = read_mt5(p, usd_inr)
        elif low.endswith(".xlsx") and not name.startswith("~$"):
            rows = read_broker(p)
        else:
            continue
        for r in rows:
            if r["trade_id"] in seen:
                if r != seen[r["trade_id"]]:
                    warnings.warn(
                        "Conflicting duplicate ID %s. Keeping the first record."
                        % r["trade_id"]
                    )
            else:
                seen[r["trade_id"]] = r

    return sorted(seen.values(), key=lambda r: (r["close_date"], r["close_time"], r["trade_id"]))


def summarize(trades):
    total = len(trades)
    win = sum(1 for t in trades if t["pnl_inr"] > 0)
    lose = sum(1 for t in trades if t["pnl_inr"] < 0)
    by = {}
    days = {}
    for t in trades:
        b = by.setdefault(t["asset_class"], {"trades": 0, "net_pnl_inr": 0.0})
        b["trades"] += 1
        b["net_pnl_inr"] += t["pnl_inr"]
        d = days.setdefault(t["close_date"], {"trades": 0, "net": 0.0})
        d["trades"] += 1
        d["net"] += t["pnl_inr"]
    daily, cum = [], 0.0
    for date in sorted(days):
        cum += days[date]["net"]
        daily.append({"date": date, "trades": days[date]["trades"],
                      "net_pnl_inr": round(days[date]["net"], 2),
                      "cumulative_pnl_inr": round(cum, 2)})
    return {"total_trades": total, "winning_trades": win, "losing_trades": lose,
            "win_rate": round(win / total, 4) if total else 0.0,
            "net_pnl_inr": round(sum(t["pnl_inr"] for t in trades), 2),
            "by_asset_class": {k: {"trades": v["trades"], "net_pnl_inr": round(v["net_pnl_inr"], 2)}
                               for k, v in sorted(by.items())},
            "daily": daily}


def find_violations(trades, rules):
    by_day = {}
    for t in trades:
        by_day.setdefault(t["close_date"], []).append(t)
    out = []
    for date in sorted(by_day):
        ts = by_day[date]
        for rule in rules:
            rtype = rule.get("type") or rule.get("rule_id") or rule.get("id")
            rid = rule.get("id") or rule.get("rule_id") or rtype
            limit = rule["limit"]
            detail = None
            if rtype == "max_trades_per_day":
                if len(ts) > limit:
                    detail = "%d trades (limit %s)" % (len(ts), limit)
            elif rtype == "max_daily_loss_inr":
                net = round(sum(t["pnl_inr"] for t in ts), 2)
                if net < -limit:
                    detail = "loss %.2f INR (limit %s)" % (-net, limit)
            elif rtype == "stop_after_consecutive_losses":
                mt5 = sorted((t for t in ts if t["close_time"]), key=lambda t: (t["close_time"], t["trade_id"]))
                streak = 0
                for t in mt5:
                    if streak >= limit:
                        detail = "trade at %s after %d consecutive losses (limit %s)" % (t["close_time"], streak, limit)
                        break
                    streak = streak + 1 if t["pnl_native"] < 0 else 0
            elif rtype == "max_lot_size":
                big = [t for t in ts if t["source"] == "mt5" and t["quantity"] > limit]
                if big:
                    detail = "lot size %s (limit %s)" % (max(t["quantity"] for t in big), limit)
            if detail:
                out.append({"date": date, "rule_id": rid, "detail": detail})
    return out


def template_report(summary, violations):
    days = summary["daily"]
    counts = {}
    for v in violations:
        counts[v["rule_id"]] = counts.get(v["rule_id"], 0) + 1
    lines = ["# Weekly Discipline Report", ""]
    if not days:
        return "\n".join(lines + ["No trades were found this week."]) + "\n"
    worst = min(days, key=lambda d: d["net_pnl_inr"])
    best = max(days, key=lambda d: d["net_pnl_inr"])
    lines.append("You closed %d trades with a %.1f%% win rate and a net result of %.2f INR."
                 % (summary["total_trades"], summary["win_rate"] * 100, summary["net_pnl_inr"]))
    lines.append("Your best day was %s (%.2f INR) and your worst was %s (%.2f INR)."
                 % (best["date"], best["net_pnl_inr"], worst["date"], worst["net_pnl_inr"]))
    if violations:
        bad_days = len({v["date"] for v in violations})
        top = max(sorted(counts), key=lambda k: counts[k])
        lines.append("You broke your own rules %d times across %d of %d trading days. "
                     "Most often broken: %s (%d times)."
                     % (len(violations), bad_days, len(days), top, counts[top]))
        lines.append("Coach's note: rules only work if you obey them when it hurts. "
                     "Next week, stop trading for the day the moment a limit is hit, "
                     "starting with %s." % top)
    else:
        lines.append("You followed every rule this week. Keep that discipline.")
    return "\n".join(lines) + "\n"


def llm_report(summary, violations):
    import urllib.request
    key = os.environ.get("GROQ_API_KEY")
    if not key:
        return None
    prompt = ("Write a weekly trading discipline report for a trader. Max 200 words, honest, specific, "
              "coach-like.\nSummary: %s\nViolations: %s" % (json.dumps(summary), json.dumps(violations)))
    body = json.dumps({"model": "llama-3.1-8b-instant", "messages": [{"role": "user", "content": prompt}]}).encode()
    req = urllib.request.Request("https://api.groq.com/openai/v1/chat/completions", data=body,
                                 headers={"Authorization": "Bearer " + key, "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=20) as r:
        text = json.load(r)["choices"][0]["message"]["content"].strip()
    return "# Weekly Discipline Report\n\n" + " ".join(text.split()[:200]) + "\n" if text else None


def write_report(summary, violations, path):
    text = None
    try:
        text = llm_report(summary, violations)
    except Exception:
        text = None
    if not text:
        text = template_report(summary, violations)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--rules", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    with open(a.rules, encoding="utf-8") as f:
        cfg = json.load(f)
    rules = cfg["rules"]
    if isinstance(rules, dict):
        rules = [{"id": k, "type": k, "limit": v} if not isinstance(v, dict) else dict({"id": k, "type": k}, **v)
                 for k, v in rules.items()]
    trades = load_trades(a.data, float(cfg["usd_inr"]))
    summary = summarize(trades)
    violations = find_violations(trades, rules)
    os.makedirs(a.out, exist_ok=True)
    with open(os.path.join(a.out, "trades.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=COLS, lineterminator="\n")
        w.writeheader()
        w.writerows(trades)
    with open(os.path.join(a.out, "summary.json"), "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2)
    with open(os.path.join(a.out, "violations.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["date", "rule_id", "detail"], lineterminator="\n")
        w.writeheader()
        w.writerows(violations)
    write_report(summary, violations, os.path.join(a.out, "report.md"))


if __name__ == "__main__":
    main()
