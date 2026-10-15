import os
import re
import sys
import importlib
from dataclasses import dataclass, field
from typing import Dict, List, Any, Optional

try:
    import yaml
except ImportError:
    yaml = None

try:
    from utils_os import obter_diretorio_executavel, obter_caminho_recurso
except ImportError:
    def obter_diretorio_executavel():
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    def obter_caminho_recurso(*caminhos):
        try:
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, *caminhos)


@dataclass
class SkillDefinition:
    name: str
    description: str
    triggers: List[str] = field(default_factory=list)
    module: str = ""
    action: str = ""
    parameters: Dict[str, Any] = field(default_factory=dict)
    doc_body: str = ""
    filepath: str = ""

    def to_tool_declaration(self) -> Dict[str, Any]:
        """Converte a definição da skill para o schema de Tool da OpenAI/OpenRouter."""
        properties = {}
        required = []

        for param_name, param_info in self.parameters.items():
            if isinstance(param_info, dict):
                prop = {
                    "type": param_info.get("type", "string"),
                    "description": param_info.get("description", "")
                }
                if "items" in param_info:
                    prop["items"] = param_info["items"]
                if "enum" in param_info:
                    prop["enum"] = param_info["enum"]
                properties[param_name] = prop

                if param_info.get("required", False):
                    required.append(param_name)
            else:
                properties[param_name] = {"type": "string", "description": str(param_info)}

        return {
            "name": self.name,
            "description": self.description,
            "parameters": {
                "type": "object",
                "properties": properties,
                "required": required
            }
        }


class GerenciadorSkills:
    """
    Gerenciador dinâmico de habilidades (Skills) baseadas em Markdown (.md).
    Permite criar, modificar e carregar skills adicionando arquivos na pasta 'skills/'
    sem necessidade de recompilar o executável.
    """

    def __init__(self, diretorios_extras: Optional[List[str]] = None):
        self.diretorios_extras = diretorios_extras or []
        self.skills: Dict[str, SkillDefinition] = {}
        self.carregar_todas_skills()

    def obter_diretorios_busca(self) -> List[str]:
        """Retorna as pastas onde arquivos .md de skills podem estar."""
        caminhos = []

        # 1. Pastas extras informadas
        for d in self.diretorios_extras:
            if d and os.path.isdir(d) and d not in caminhos:
                caminhos.append(d)

        # 2. Pasta 'skills' ao lado do executável ou raiz do projeto
        try:
            d_exe = os.path.join(obter_diretorio_executavel(), "skills")
            if os.path.isdir(d_exe) and d_exe not in caminhos:
                caminhos.append(d_exe)
        except Exception:
            pass

        # 3. Pasta 'skills' embutida no pacote (sys._MEIPASS)
        try:
            d_rec = obter_caminho_recurso("skills")
            if os.path.isdir(d_rec) and d_rec not in caminhos:
                caminhos.append(d_rec)
        except Exception:
            pass

        # 4. Pasta de trabalho atual
        d_cwd = os.path.join(os.getcwd(), "skills")
        if os.path.isdir(d_cwd) and d_cwd not in caminhos:
            caminhos.append(d_cwd)

        return caminhos

    def parse_skill_markdown(self, conteudo: str, filepath: str = "") -> Optional[SkillDefinition]:
        """Faz o parsing do Frontmatter YAML e do corpo em Markdown."""
        conteudo = conteudo.strip()
        if not conteudo.startswith("---"):
            return None

        partes = re.split(r"^---\s*$", conteudo, maxsplit=2, flags=re.MULTILINE)
        if len(partes) < 3:
            return None

        frontmatter_raw = partes[1].strip()
        doc_body = partes[2].strip()

        meta = {}
        if yaml:
            try:
                meta = yaml.safe_load(frontmatter_raw) or {}
            except Exception as e:
                print(f"[SKILLS] Aviso: Erro ao parsear YAML em {filepath}: {e}")
                meta = self._parse_yaml_fallback(frontmatter_raw)
        else:
            meta = self._parse_yaml_fallback(frontmatter_raw)

        if not isinstance(meta, dict):
            return None

        name = meta.get("name")
        if not name:
            # Fallback para o nome do arquivo sem extensão
            name = os.path.splitext(os.path.basename(filepath))[0] if filepath else "skill_desconhecida"

        description = meta.get("description", "")
        triggers = meta.get("triggers", [])
        if isinstance(triggers, str):
            triggers = [t.strip() for t in triggers.split(",") if t.strip()]
        elif not isinstance(triggers, list):
            triggers = []

        module = str(meta.get("module", ""))
        action = str(meta.get("action", ""))
        parameters = meta.get("parameters", {})
        if not isinstance(parameters, dict):
            parameters = {}

        return SkillDefinition(
            name=name,
            description=description,
            triggers=triggers,
            module=module,
            action=action,
            parameters=parameters,
            doc_body=doc_body,
            filepath=filepath
        )

    def _parse_yaml_fallback(self, texto: str) -> Dict[str, Any]:
        """Parser básico de fallback caso o PyYAML não esteja disponível."""
        resultado = {}
        chave_atual = None
        modo_lista = False

        for linha in texto.splitlines():
            linha_strip = linha.strip()
            if not linha_strip or linha_strip.startswith("#"):
                continue

            if linha_strip.startswith("- ") and chave_atual:
                item = linha_strip[2:].strip().strip("\"'")
                if not isinstance(resultado.get(chave_atual), list):
                    resultado[chave_atual] = []
                resultado[chave_atual].append(item)
                continue

            if ":" in linha:
                k, v = linha.split(":", 1)
                k = k.strip()
                v = v.strip().strip("\"'")
                if v:
                    resultado[k] = v
                    chave_atual = k
                    modo_lista = False
                else:
                    resultado[k] = []
                    chave_atual = k
                    modo_lista = True

        return resultado

    def carregar_todas_skills(self) -> Dict[str, SkillDefinition]:
        """Varre todas as pastas de busca e carrega arquivos .md."""
        pastas = self.obter_diretorios_busca()
        skills_carregadas = {}

        for pasta in pastas:
            try:
                for arquivo in os.listdir(pasta):
                    if arquivo.lower().endswith(".md"):
                        caminho_completo = os.path.join(pasta, arquivo)
                        try:
                            with open(caminho_completo, "r", encoding="utf-8") as f:
                                conteudo = f.read()
                            skill = self.parse_skill_markdown(conteudo, caminho_completo)
                            if skill:
                                # Sobrescreve mantendo prioridade (arquivos locais sobrescrevem embutidos)
                                skills_carregadas[skill.name] = skill
                        except Exception as e:
                            print(f"[SKILLS] Erro ao ler skill '{caminho_completo}': {e}")
            except Exception as e:
                print(f"[SKILLS] Erro ao listar pasta '{pasta}': {e}")

        self.skills = skills_carregadas
        return self.skills

    def obter_skill(self, nome: str) -> Optional[SkillDefinition]:
        return self.skills.get(nome)

    def listar_skills(self) -> List[SkillDefinition]:
        return list(self.skills.values())

    def gerar_declaracoes_tools_openrouter(self) -> List[Dict[str, Any]]:
        """Gera lista de declarações de funções para ser injetada no catálogo de Tools do chatbot."""
        return [skill.to_tool_declaration() for skill in self.skills.values()]

    def buscar_por_gatilho(self, texto: str) -> Optional[SkillDefinition]:
        """
        Encontra uma skill correspondente ao texto recebido usando correspondência
        de triggers e palavras-chave.
        """
        texto_lower = texto.lower()
        # Busca exata ou por substring em ordem decrescente de comprimento de gatilho
        melhor_skill = None
        maior_gatilho_len = 0

        for skill in self.skills.values():
            for g in skill.triggers:
                g_clean = str(g).lower().strip()
                if not g_clean:
                    continue
                # Verifica correspondência de palavra inteira ou expressão
                padrao = r"\b" + re.escape(g_clean) + r"\b"
                if re.search(padrao, texto_lower):
                    if len(g_clean) > maior_gatilho_len:
                        maior_gatilho_len = len(g_clean)
                        melhor_skill = skill

        return melhor_skill

    def executar_skill(self, nome: str, args: Dict[str, Any], funcoes_disponiveis: Optional[Dict[str, Any]] = None) -> Any:
        """
        Executa a ação da skill, seja chamando o dicionário funcoes_disponiveis
        ou importando dinamicamente o módulo e executando a função.
        """
        # 1. Se já existir no dicionário de funções do app (conecta_hub2)
        if funcoes_disponiveis and nome in funcoes_disponiveis:
            func = funcoes_disponiveis[nome]
            try:
                return func(**args)
            except TypeError:
                try:
                    return func()
                except Exception as e:
                    return {"erro": f"Erro de assinatura ao executar {nome}: {e}"}
            except Exception as e:
                return {"erro": str(e)}

        skill = self.obter_skill(nome)
        if not skill:
            return {"erro": f"Skill '{nome}' não encontrada."}

        # 2. Se a skill aponta para uma função no funcoes_disponiveis com outro nome de action
        if funcoes_disponiveis and skill.action in funcoes_disponiveis:
            func = funcoes_disponiveis[skill.action]
            try:
                return func(**args)
            except TypeError:
                try:
                    return func()
                except Exception as e:
                    return {"erro": f"Erro de assinatura ao executar {skill.action}: {e}"}
            except Exception as e:
                return {"erro": str(e)}

        # 3. Importação dinâmica do módulo e execução da função
        if skill.module and skill.action:
            try:
                mod = importlib.import_module(skill.module)
                func = getattr(mod, skill.action, None)
                if not func:
                    return {"erro": f"Ação '{skill.action}' não encontrada no módulo '{skill.module}'."}
                return func(**args)
            except Exception as e:
                return {"erro": f"Erro ao executar {skill.module}.{skill.action}: {e}"}

        return {"erro": f"A skill '{nome}' não possui uma ação ou módulo configurado para execução."}


# Instância global singleton do gerenciador para fácil acesso
_gerenciador_global = None

def obter_gerenciador_skills() -> GerenciadorSkills:
    global _gerenciador_global
    if _gerenciador_global is None:
        _gerenciador_global = GerenciadorSkills()
    return _gerenciador_global
