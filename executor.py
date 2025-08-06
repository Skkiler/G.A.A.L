#executor.py
import os, re, shutil, zipfile, tempfile
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd

def match_filters(f, filters):
    if not filters:
        return True
    # --- Por extensão ---
    ext = f.suffix.lower()[1:]  # ".pdf" -> "pdf"
    if 'types' in filters and filters['types']:
        if ext not in filters['types']:
            return False
    # --- Por data ---
    stat = f.stat()
    if 'date_start' in filters and filters['date_start']:
        try:
            dt_ini = datetime.strptime(filters['date_start'], '%Y-%m-%d')
        except Exception:
            dt_ini = datetime.strptime(filters['date_start'], '%d-%m-%Y')
        if datetime.fromtimestamp(stat.st_mtime) < dt_ini:
            return False
    if 'date_end' in filters and filters['date_end']:
        try:
            dt_fim = datetime.strptime(filters['date_end'], '%Y-%m-%d')
        except Exception:
            dt_fim = datetime.strptime(filters['date_end'], '%d-%m-%Y')
        if datetime.fromtimestamp(stat.st_mtime) > dt_fim:
            return False
    return True

def buscar_subpasta(destino, nome, recursivo=True):
    destino = Path(destino)
    if not destino.exists():
        return []
    matches = []
    if recursivo:
        for p in destino.rglob("*"):
            if p.is_dir() and p.name == nome:
                matches.append(p)
    else:
        for p in destino.iterdir():
            if p.is_dir() and p.name == nome:
                matches.append(p)
    return matches

class BooleanConditionEngine:
    
    def __init__(self, names, expr):
        self.names = names
        self.expr = expr

    def evaluate(self, md, filename=""):
        e = self.expr
        fname = filename.lower()
        # 1) trata literais fixos em "entre aspas"
        for lit in re.findall(r'"([^"]+)"', e):
            found = lit.lower() in fname
            e = e.replace(f'"{lit}"', str(found))
        # 2) substitui !Cond! → True/False
        for k, v in md.items():
            e = e.replace(f"!{k}!", str(bool(v)))
        # 3) negações {Cond}
        def neg(m):
            nome = m.group(1)
            return "not " + str(bool(md.get(nome, False)))
        e = re.sub(r"\{([A-Za-z0-9_]+)\}", neg, e)
        # 4) operadores lógicos
        e = e.replace('&', ' and ').replace('|', ' or ')
        try:
            return eval(e)
        except Exception:
            return False

class ExcelConditionEngine:
    
    def __init__(self, path, cols, prims, expr):
        self.path = Path(path)
        self.cols = cols
        self.boolean = BooleanConditionEngine(list(cols.keys()), expr)
        self.principais = list(prims)
        self._load()

    def _load(self):
        try:
            tmp = Path(tempfile.mkdtemp()) / self.path.name
            shutil.copy2(self.path, tmp)
            self.df = pd.read_excel(tmp, dtype=str)
        except Exception:
            self.df = None
        finally:
            shutil.rmtree(tmp.parent, ignore_errors=True)

    def evaluate(self, filename):
        if self.df is None or len(self.df) == 0:
            return False
        filename_lower = filename.lower()
        for idx, row in self.df.iterrows():
            md = {}
            for n, col in self.cols.items():
                try:
                    val = str(row[col]).strip() if col in row else ""
                    md[n] = val.lower() in filename_lower if val else False
                except Exception:
                    md[n] = False
            if self.boolean.evaluate(md, filename_lower):
                return True
        return False

    def get_principais_values(self, filename, sep="_"):
        if self.df is None or len(self.df) == 0:
            return None
        filename_lower = filename.lower()
        for idx, row in self.df.iterrows():
            md = {}
            for n, col in self.cols.items():
                try:
                    val = str(row[col]).strip() if col in row else ""
                    md[n] = val.lower() in filename_lower if val else False
                except Exception:
                    md[n] = False
            if self.boolean.evaluate(md, filename_lower):
                vals = []
                for n in self.principais:
                    col = self.cols.get(n, None)
                    if col is not None and col in row:
                        vals.append(str(row[col]).strip())
                return sep.join(vals) if vals else None
        return None

    def all_matching_rows(self, filename):
        """
        Retorna todas as linhas do excel que são compatíveis com o arquivo (baseado na expressão e colunas)
        """
        if self.df is None or len(self.df) == 0:
            return []
        filename_lower = filename.lower()
        matches = []
        for idx, row in self.df.iterrows():
            md = {}
            for n, col in self.cols.items():
                try:
                    val = str(row[col]).strip() if col in row else ""
                    md[n] = val.lower() in filename_lower if val else False
                except Exception:
                    md[n] = False
            if self.boolean.evaluate(md, filename_lower):
                matches.append(row)
        return matches

class FolderConditionEngine:
    
    def __init__(self, base, cols, prims, sep, expr):
        self.base = Path(base)
        self.cols = cols
        self.boolean = BooleanConditionEngine(list(cols.keys()), expr)
        self.sep = sep
        self.subs = [d.name for d in self.base.iterdir() if d.is_dir()]
        self.principais = list(prims)

    def matched_subfolders(self, filename):
        out = []
        filename_lower = filename.lower()
        for sub_name in self.subs:
            tok = sub_name.split(self.sep)
            md = {}
            for c, i in self.cols.items():
                idx = i - 1
                val = tok[idx] if (0 <= idx < len(tok)) else ""
                md[c] = (val != "" and val.lower() in filename_lower)
            if self.boolean.evaluate(md, filename_lower):
                out.append(sub_name)
        return out

    def build_principais_subfolder(self, filename):
        # Nova função para montar subpasta baseada em valores das principais
        tokens = filename.split(self.sep)
        vals = []
        for c in self.principais:
            idx = self.cols.get(c, None)
            if idx is not None and (idx-1) < len(tokens):
                vals.append(tokens[idx-1])
        return self.sep.join(vals) if vals else None

class FileManager:
    
    def __init__(self, origins, destino, action, extract_zips=False, recursivo=True, hierarchy=False, criar_subpasta=False, copy_dirs=False):
        self.origins = [Path(p) for p in origins]
        self.destino = Path(destino)
        self.action = action
        self.extract_zips = extract_zips
        self.recursivo = recursivo
        self.hierarchy = hierarchy
        self.criar_subpasta = criar_subpasta
        self.copy_dirs = copy_dirs
        self._temp_dirs = []

    def collect_files(self):
        files = set()
        for p in self.origins:
            if p.is_dir():
                # Pega tudo dentro da pasta
                found = list(p.rglob("*")) if self.recursivo else list(p.glob("*"))
                for f in found:
                    if self.extract_zips and f.is_file() and f.suffix.lower() == ".zip":
                        # Extrai todos os ZIPs encontrados dentro da pasta
                        tmp_dir = Path(tempfile.mkdtemp())
                        self._temp_dirs.append(tmp_dir)
                        with zipfile.ZipFile(f, 'r') as zip_ref:
                            zip_ref.extractall(tmp_dir)
                        for zf in tmp_dir.rglob("*"):
                            if zf.is_file():
                                files.add(zf)
                    elif f.is_file():
                        files.add(f)
                    elif self.copy_dirs and f.is_dir():
                        files.add(f)
            elif p.is_file() and p.suffix.lower() == ".zip" and self.extract_zips:
                # Se origem for um ZIP direto, extrai também
                tmp_dir = Path(tempfile.mkdtemp())
                self._temp_dirs.append(tmp_dir)
                with zipfile.ZipFile(p, 'r') as zip_ref:
                    zip_ref.extractall(tmp_dir)
                for zf in tmp_dir.rglob("*"):
                    if zf.is_file():
                        files.add(zf)
            elif p.is_file():
                files.add(p)
        return list(files)

    def cleanup(self):
        for d in self._temp_dirs:
            shutil.rmtree(d, ignore_errors=True)

    def process_file(self, src, sub=None, hierarchy_path=None):
        destino_final = ""
        try:
            if self.action == "delete":
                destino_final = "DELETADO"
                if src.is_dir():
                    shutil.rmtree(src)
                else:
                    src.unlink()
                return destino_final
            dst_dir = self.destino
            if hierarchy_path:
                dst_dir = dst_dir / hierarchy_path
            if sub:
                dst_dir = dst_dir / sub
            dst_dir.mkdir(parents=True, exist_ok=True)
            if src.is_file():
                shutil.copy2(src, dst_dir / src.name)
                destino_final = str(dst_dir / src.name)
            elif src.is_dir():
                # Copiar pasta (inclusive vazia)
                dest_folder = dst_dir / src.name
                if not dest_folder.exists():
                    dest_folder.mkdir(parents=True, exist_ok=True)
                for file in src.rglob("*"):
                    if file.is_file():
                        rel = file.relative_to(src)
                        (dest_folder / rel.parent).mkdir(parents=True, exist_ok=True)
                        shutil.copy2(file, dest_folder / rel)
                destino_final = str(dest_folder)
        except PermissionError:
            destino_final = "ERRO_PERMISSAO"
        return destino_final

class Executor:
    
    def __init__(self, cfg, max_workers=4, progress_callback=None, error_callback=None, complete_callback=None,
                 cancel_checker=None, report_callback=None):
        self.cfg = cfg
        self.progress = progress_callback or (lambda *a: None)
        self.error = error_callback or (lambda *a: None)
        self.complete = complete_callback or (lambda: None)
        self.cancel_checker = cancel_checker or (lambda: False)
        self.report_callback = report_callback or (lambda item: None)
        self.origins = [Path(p) for p in cfg["origens"]]
        self.fm = FileManager(
            self.origins,
            cfg["destino"],
            cfg["action"],
            extract_zips=cfg.get("extract_zips", False),
            recursivo=cfg.get("recursivo", True),
            hierarchy=cfg.get("hierarchy", False),
            criar_subpasta=cfg.get("criar_subpasta", False),
            copy_dirs=cfg.get("copy_dirs", False)
        )
        self.use_cond = cfg.get("use_conditions", True)
        self.sep = cfg.get("cond_sep", "_")
        if self.use_cond:
            if cfg["condition_mode"] == "excel":
                self.ce = ExcelConditionEngine(cfg["excel"], cfg["colunas"], cfg["principais"], cfg["condition_expression"])
            else:
                self.ce = FolderConditionEngine(cfg["cond_folder"], cfg["colunas"], cfg["principais"], self.sep, cfg["condition_expression"])
        else:
            self.ce = None
        self.max_workers = max_workers
        self.sobra = cfg.get("sobra", None)
        self.zip_dest = cfg.get("zip_dest", False)
        self.sobra_enabled = cfg.get("sobra_enabled", False)
        
        # ==== configuração de renomeação ====
        rename_cfg = cfg.get("rename", {})
        self.rename_enabled = rename_cfg.get("enabled", False)
        self.rename_pattern = rename_cfg.get("pattern", "")

    def get_sobra_path(self):
        """
        Retorna o caminho absoluto para salvar sobras:
        - Se self.sobra vazio, nulo ou só espaços, usa '[destino]/sobra'
        - Se self.sobra for só nome (sem barra), salva como subpasta do destino
        - Se for caminho absoluto/relativo, usa direto
        """
        sobra = self.sobra
        if not sobra or not sobra.strip():
            sobra = "sobra"
        p = Path(sobra)
        if not p.is_absolute() and not (p.parts[0] == "." or p.parts[0] == ".."):
            return str(self.fm.destino / sobra)
        else:
            return str(p)

    def run(self):
        files = self.fm.collect_files()
        total = len(files)
        done = 0
        reports = []
        try:
            with ThreadPoolExecutor(max_workers=self.cfg["max_workers"]) as pool:
                futures = {pool.submit(self._process, f): f for f in files}
                for fut in as_completed(futures):
                    if self.cancel_checker():
                        break
                    f = futures[fut]
                    try:
                        res = fut.result()
                        if isinstance(res, list):
                            for r in res:
                                self.report_callback(r)
                                reports.append(r)
                        elif res:
                            self.report_callback(res)
                            reports.append(res)
                    except Exception as e:
                        self.error(str(e))
                    finally:
                        done += 1
                        self.progress(done, total)
        finally:
            self.fm.cleanup()
            if self.zip_dest:
                self._zip_destination()
            self.complete()

    def _process(self, f):
        # 1) filtrar por extensão/data
        if not match_filters(f, self.cfg.get("file_filters", {})):
            return None

        # 2) hierarquia física (quando hierarchy=True e criar_subpasta=False)
        rel_hierarchy = None
        if not self.cfg.get("criar_subpasta", False) and self.fm.hierarchy:
            for origem in self.origins:
                try:
                    rel = f.relative_to(origem).parent
                    rel_hierarchy = rel if rel != Path('.') else None
                    break
                except Exception:
                    continue

        # helper: aplica delete ou copy com renomeação, hierarquia e subpasta
        def _do_transfer(src, sub=None, hierarchy_path=None):
            # delete
            if self.cfg["action"] == "delete":
                if src.is_dir():
                    shutil.rmtree(src)
                else:
                    src.unlink()
                return {
                    "arquivo": src.name,
                    "origem": str(src),
                    "destino": "DELETADO",
                    "acao": "delete"
                }
            # monta pasta destino
            dst_dir = self.fm.destino
            if hierarchy_path:
                dst_dir = dst_dir / hierarchy_path
            if sub:
                dst_dir = dst_dir / sub
            dst_dir.mkdir(parents=True, exist_ok=True)

            # renomeação
            final_name = src.name
            if self.rename_enabled:
                pattern = self.rename_pattern
                for lit in re.findall(r'"([^"]+)"', pattern):
                    pattern = pattern.replace(f'"{lit}"', lit)
                # variáveis de condição !Cond!
                for var in re.findall(r'!([^!]+)!', pattern):
                    val = ""
                    if hasattr(self.ce, "get_principais_values"):
                        val = self.ce.get_principais_values(src.stem, self.sep) or ""
                    pattern = pattern.replace(f'!{var}!', val)
                final_name = pattern + src.suffix

            # copy file ou pasta
            report = None
            if src.is_file():
                shutil.copy2(src, dst_dir / final_name)
                report = {
                    "arquivo": src.name,
                    "origem": str(src),
                    "destino": str(dst_dir / final_name),
                    "acao": self.cfg["action"]
                }
            else:
                dest_folder = dst_dir / src.name
                dest_folder.mkdir(parents=True, exist_ok=True)
                for item in src.rglob("*"):
                    if item.is_file():
                        rel = item.relative_to(src)
                        (dest_folder / rel.parent).mkdir(parents=True, exist_ok=True)
                        shutil.copy2(item, dest_folder / rel)
                report = {
                    "arquivo": src.name,
                    "origem": str(src),
                    "destino": str(dest_folder),
                    "acao": self.cfg["action"]
                }
            return report

        # 2.1) se use_cond=True mas expressão vazia, aceitar todos os arquivos
        expr = self.cfg.get("condition_expression", "").strip()
        if self.use_cond and not expr:
            return _do_transfer(f, hierarchy_path=rel_hierarchy)

        # 2.2) expressão isolada sem nenhuma coluna cadastrada
        if self.use_cond and expr and not self.ce.boolean.names:
            # avalia literais em "" + operadores lógicos
            if self.ce.boolean.evaluate({}, f.stem):
                return _do_transfer(f, hierarchy_path=rel_hierarchy)
            else:
                return None

        # 3) sem condições ativas => transfere direto
        if not self.use_cond:
            return _do_transfer(f, hierarchy_path=rel_hierarchy)

        # 4) FolderConditionEngine
        if isinstance(self.ce, FolderConditionEngine):
            subpasta = None
            if self.cfg["principais"]:
                subpasta = self.ce.build_principais_subfolder(f.stem)
            find_sub = self.cfg.get("find_subpasta", False)
            criar_sub = self.cfg.get("criar_subpasta", False)
            multipl = self.cfg.get("multiply", False)
            tem_sobra = self.sobra_enabled and bool(self.sobra)

            # procurar subpasta existente
            if find_sub and subpasta:
                encontrados = buscar_subpasta(self.cfg["destino"], subpasta, self.cfg.get("recursivo", True))
                if encontrados:
                    if multipl:
                        return [ _do_transfer(f, None, p.relative_to(self.fm.destino)) for p in encontrados ]
                    return _do_transfer(f, None, encontrados[0].relative_to(self.fm.destino))
                if criar_sub:
                    return _do_transfer(f, subpasta)
                if tem_sobra:
                    return _do_transfer(f, self.get_sobra_path(), rel_hierarchy)
                return None

            # criar subpasta baseada em principais + hierarquia
            if criar_sub and self.cfg.get("hierarchy", False) and self.cfg["principais"]:
                tokens = f.stem.split(self.sep)
                parts = [tokens[self.cfg["colunas"][c]-1]
                         for c in self.cfg["principais"]
                         if 0 <= self.cfg["colunas"][c]-1 < len(tokens)]
                return _do_transfer(f, hierarchy_path=Path(*parts) if parts else None)
            if criar_sub and self.cfg["principais"]:
                return _do_transfer(f, subpasta)

            # match em subpastas pela expressão
            matches = self.ce.matched_subfolders(f.stem)
            if multipl:
                reports = [ _do_transfer(f, sub, rel_hierarchy) for sub in matches ]
                if not reports and tem_sobra:
                    reports.append(_do_transfer(f, self.get_sobra_path(), rel_hierarchy))
                return reports or None
            if matches:
                return _do_transfer(f, matches[0], rel_hierarchy)
            if tem_sobra:
                return _do_transfer(f, self.get_sobra_path(), rel_hierarchy)
            return None

        # 5) ExcelConditionEngine
        if isinstance(self.ce, ExcelConditionEngine):
            find_sub = self.cfg.get("find_subpasta", False)
            criar_sub = self.cfg.get("criar_subpasta", False)
            multipl = self.cfg.get("multiply", False)
            tem_sobra = self.sobra_enabled and bool(self.sobra)
            matched = []

            # buscar linhas que batem
            if self.ce.df is not None:
                fname = f.stem.lower()
                for _, row in self.ce.df.iterrows():
                    md = { n: (str(row[col]).strip().lower() in fname) for n, col in self.ce.cols.items() if row[col] }
                    if self.ce.boolean.evaluate(md, fname):
                        matched.append(row)

            # múltiplos
            if multipl and matched:
                reports = []
                for row in matched:
                    vals = [str(row[self.ce.cols[n]]).strip() for n in self.cfg["principais"]]
                    subp = self.sep.join(vals)
                    if find_sub and subp:
                        enc = buscar_subpasta(self.cfg["destino"], subp, self.cfg.get("recursivo", True))
                        for pasta in enc:
                            reports.append(_do_transfer(f, None, pasta.relative_to(self.fm.destino)))
                        continue
                    if criar_sub and subp:
                        reports.append(_do_transfer(f, subp))
                        continue
                    reports.append(_do_transfer(f, None, rel_hierarchy))
                return reports or None

            # único match
            if matched:
                row = matched[0]
                vals = [str(row[self.ce.cols[n]]).strip() for n in self.cfg["principais"]]
                subp = self.sep.join(vals)
                if find_sub and subp:
                    enc = buscar_subpasta(self.cfg["destino"], subp, self.cfg.get("recursivo", True))
                    if enc:
                        return _do_transfer(f, None, enc[0].relative_to(self.fm.destino))
                if criar_sub and subp:
                    return _do_transfer(f, subp)
                return _do_transfer(f, None, rel_hierarchy)

            # sobra
            if tem_sobra:
                return _do_transfer(f, self.get_sobra_path(), rel_hierarchy)
            return None

    def _zip_destination(self):
        dest_dir = Path(self.cfg["destino"])
        zip_path = dest_dir.parent / (dest_dir.name + ".zip")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(dest_dir):
                for file in files:
                    absfile = Path(root) / file
                    relpath = absfile.relative_to(dest_dir.parent)
                    zipf.write(absfile, relpath)
