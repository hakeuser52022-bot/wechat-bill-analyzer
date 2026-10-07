"""开发检查与报名检查分开；自动扫描不替代人工隐私审查。"""
import argparse,json,hashlib,re,subprocess,sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8',errors='replace')
ROOT=Path(__file__).resolve().parents[1]
def main():
    ap=argparse.ArgumentParser();ap.add_argument('--submission',action='store_true');ap.add_argument('--run-code',action='store_true');a=ap.parse_args()
    failures=0
    for script in ['tests/test_core.py','tools/check_links.py','tools/verify_numbers.py','tools/check_food_consistency.py']:
        p=subprocess.run([sys.executable,'-X','utf8',script],cwd=ROOT,capture_output=True,text=True,encoding='utf-8')
        print(('PASS ' if p.returncode==0 else 'FAIL ')+script)
        if p.returncode: print(p.stdout+p.stderr);failures+=1
    # Packaged artifact hashes detect unintended replacement; they do not certify privacy.
    manifest=ROOT/'examples/artifacts.sha256.json'
    if not manifest.exists(): print('FAIL 缺少演示产物清单');failures+=1
    else:
        hashes=json.loads(manifest.read_text(encoding='utf-8'))
        for rel,h in hashes.items():
            path=ROOT/rel
            if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest()!=h:
                print('FAIL 演示产物变更，请重跑 run_demo.py：'+rel);failures+=1
        for path in ROOT.rglob('*'):
            if path.suffix.lower() in ('.xlsx','.xls','.csv') and path.relative_to(ROOT).as_posix() not in hashes:
                print('FAIL 未登记的数据文件：'+path.name);failures+=1
    pending=[]
    for rel in ['docs/intro.md','LICENSE']:
        path=ROOT/rel
        if not path.exists():pending.append(rel+' 缺失');continue
        txt=path.read_text(encoding='utf-8')
        if re.search(r'【[^】]+】|<[^>]+>|你的用户名',txt):pending.append(rel)
    try:
        from docx import Document
        d=Document(ROOT/'docs/个人简介（待填写）.docx')
        txt=' '.join(p.text for p in d.paragraphs)
        if re.search(r'【[^】]+】|你的用户名',txt):pending.append('Word 个人简介')
    except Exception as e: print('FAIL Word 检查失败：'+str(e));failures+=1
    if pending:
        print(('FAIL ' if a.submission else '待填写 ')+'、'.join(pending))
        if a.submission:failures+=1
    print('请人工核对个人信息、文件内容、仓库历史及原招新通知；自动检查不证明无隐私风险。')
    print('结果：'+('检查失败' if failures else ('报名自动检查通过，仍需人工终检' if a.submission else '开发检查通过；不代表材料可以直接提交')))
    return 1 if failures else 0
if __name__=='__main__':sys.exit(main())
