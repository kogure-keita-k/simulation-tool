# -*- coding: utf-8 -*-
"""型式から冷房能力(kW)を推定する。完全ではないため『推定』フラグを返す。
ルール: 三菱/ダイキン/日立の業務用は型式中の3桁数字÷10 (280→28.0, 224→22.4, 112→11.2, 80→8.0)。
東芝(ROA/ROB/RCR系)は数字÷100 (1125→11.2, 633→6.3, 565→5.6)。
ルームエアコン(MUZ/MUCZ/MSZ/CU/RAS-無印)は先頭2桁÷10 (36→3.6, 56→5.6)。
判定不能・不明はNoneを返す。"""
import re
def decode(model, maker=""):
    if not model: return None, "型式なし→要確認"
    m=str(model); mk=str(maker or "")
    if any(x in m for x in ["不明","読み取り","-"]) and not re.search(r"\d{2,}",m):
        return None, "型式不明→要確認"
    nums=re.findall(r"\d{2,4}", m)
    if not nums: return None, "数字なし→要確認"
    is_toshiba = ("東芝" in mk) or ("toshiba" in mk.lower()) or bool(re.match(r"^(ROA|ROB)", m))
    is_room = bool(re.match(r"^(MUZ|MUCZ|MSZ|MFZ|CU|AU)", m)) or ("ルーム" in m)
    tok = max(nums, key=lambda x:(len(x), int(x)))
    v=int(tok); flag="推定"
    if is_toshiba:
        cap=int(tok[:-1])/10 if len(tok)>=3 else v/10
    elif is_room:
        cap=int(tok[:2])/10
    else:
        cap=v/10 if len(tok)<=3 else v/100
    if not (1.5<=cap<=60):
        for alt in (v/10, v/100, int(tok[:2])/10):
            if 1.5<=alt<=60: cap=alt; break
    if not (1.5<=cap<=60): return None, "推定値が範囲外→要確認"
    return round(cap,1), flag

if __name__=="__main__":
    tests=[("PUZ-ERMP280KA4","三菱電機",28.0),("PUZ-ZRMP80HA14","三菱",8.0),("PUHV-P500SDM-E","三菱",50.0),
    ("RZRP224A","ダイキン",22.4),("PURY-P280DMG5-BSG","三菱",28.0),("RAS-AP80SH2","日立",8.0),
    ("RCR-AP280HVG","日立",28.0),("ROA-AP1125HS","東芝",11.2),("ROA-633H","東芝",6.3),("ROA-AP565HSJZ1","東芝",5.6),
    ("PUZ-ERMP112LA4","三菱",11.2),("PUHY-RP224DMG6","三菱",22.4),("MUCZ-G5617S","三菱",5.6),("CU-XS360D2","パナソニック",3.6),
    ("PUZ-ERMP160LA2","三菱",16.0),("PUHV-P335SDM-E","三菱",33.5),("RAS-J140H1","日立",14.0)]
    ok=0
    for md,mk,exp in tests:
        cap,flag=decode(md,mk); mark="OK" if cap==exp else "✗"
        if cap==exp: ok+=1
        print(f"{mark} {md:<22}{mk:<8} → {cap} (期待{exp}) {flag}")
    print(f"\n精度: {ok}/{len(tests)}")
