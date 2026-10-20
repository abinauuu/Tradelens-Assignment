import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import run


def test_asset_class_and_symbol():
    assert run.clean_symbol("xauusd.m") == "XAUUSD"
    assert run.asset_class("XAUUSD") == "metals"
    assert run.asset_class("BANKNIFTY26SEP52000PE") == "index_options"
    assert run.asset_class("FOO") == "other"


def test_mt5_worked_example(tmp_path):
    p = tmp_path / "a.csv"
    p.write_text("Ticket,Open Time,Type,Volume,Symbol,Open Price,Close Time,Close Price,Commission,Swap,Profit\n"
                 "50012345,2026.09.08 16:00:00,buy,1.0,XAUUSD.m,1,2026.09.08 19:31:38,1,-7.0,0.8,-142.63\n"
                 "1,2026.09.08 16:00:00,balance,,,,2026.09.08 16:00:00,,0,0,5000\n")
    rows = run.read_mt5(str(p), 84.0)
    assert len(rows) == 1
    r = rows[0]
    assert (r["close_date"], r["close_time"], r["pnl_native"], r["pnl_inr"]) == ("2026-09-09", "01:01:38", -148.83, -12501.72)


def test_sample_dedup_and_broker():
    base = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "sample", "data")
    trades = run.load_trades(base, 84.0)
    ids = [t["trade_id"] for t in trades]
    assert len(ids) == len(set(ids))
    b = [t for t in trades if t["source"] == "broker_in"]
    assert len(b) == 15
    assert any(t["pnl_native"] == -30432.67 for t in b)


def test_consecutive_losses():
    mk = lambda tm, p: {"close_date": "d", "close_time": tm, "pnl_native": p, "pnl_inr": p, "source": "mt5", "quantity": 1, "trade_id": tm}
    rules = [{"id": "c", "type": "stop_after_consecutive_losses", "limit": 2}]
    assert run.find_violations([mk("10:00:00", -1), mk("11:00:00", -1), mk("12:00:00", 5)], rules)
    assert not run.find_violations([mk("10:00:00", -1), mk("11:00:00", 1), mk("12:00:00", -1)], rules)


def _make_xlsx(path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Statement"])
    ws.append([])
    ws.append([])
    ws.append(["Order ID", "Symbol", "Trade Date", "Quantity", "Buy Value", "Sell Value", "Realized P&L"])
    ws.append([1103456, "NIFTY26SEP25100CE", "15-09-2026", 75, 1, 2, "-3,200.00"])
    ws.append([None] * 7)
    ws.append([1103457, "banknifty26sep1ce", "16-09-2026", 60, 1, 2, "12,450.50"])
    ws.append(["Total", None, None, None, None, None, 9250.5])
    wb.save(path)


def test_broker_messy_excel(tmp_path):
    p = tmp_path / "b.xlsx"
    _make_xlsx(str(p))
    rows = run.read_broker(str(p))
    assert [r["trade_id"] for r in rows] == ["in-1103456", "in-1103457"]
    assert rows[0]["pnl_inr"] == -3200.0 and rows[0]["close_time"] == "" and rows[0]["quantity"] == 75
    assert rows[1]["pnl_native"] == 12450.5 and rows[1]["asset_class"] == "index_options"


def test_summary_and_rules():
    mk = lambda i, d, p, q=1: {"trade_id": str(i), "source": "mt5", "asset_class": "metals", "close_date": d,
                               "close_time": "10:0%d:00" % i, "pnl_native": p, "pnl_inr": p, "quantity": q}
    trades = [mk(1, "2026-09-07", -100), mk(2, "2026-09-07", 50, 3.0), mk(3, "2026-09-08", 20)]
    s = run.summarize(trades)
    assert (s["total_trades"], s["winning_trades"], s["losing_trades"], s["win_rate"]) == (3, 2, 1, 0.6667)
    assert s["daily"][1]["cumulative_pnl_inr"] == -30.0
    rules = [{"id": "R1", "type": "max_trades_per_day", "limit": 1},
             {"id": "R2", "type": "max_daily_loss_inr", "limit": 10},
             {"id": "R4", "type": "max_lot_size", "limit": 2.0}]
    got = {(v["date"], v["rule_id"]) for v in run.find_violations(trades, rules)}
    assert got == {("2026-09-07", "R1"), ("2026-09-07", "R2"), ("2026-09-07", "R4")}
