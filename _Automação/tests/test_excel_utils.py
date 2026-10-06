from datetime import datetime

from treinamentos_its.excel_utils import (
    formatar_data_texto,
    normalize_name,
    normalizar_versao,
    parse_data,
    sem_acentos,
    versao_e_numerica,
)


# --- sem_acentos (B6) --------------------------------------------------------


def test_sem_acentos_tira_acento_mantendo_caixa_e_espacos():
    assert sem_acentos("João") == "Joao"
    assert sem_acentos("Estagiário") == "Estagiario"
    assert sem_acentos("  Ática ") == "  Atica "  # não mexe em bordas


def test_sem_acentos_cedilha_e_varios_diacriticos():
    assert sem_acentos("Conclusão çedilha ÀÉÎÕÜ") == "Conclusao cedilha AEIOU"


def test_sem_acentos_texto_sem_acento_passa_igual():
    assert sem_acentos("ANALISTA ADM") == "ANALISTA ADM"


# --- normalize_name (B6: agora ignora acento) -------------------------------


def test_normalize_name_sobe_caixa_e_tira_acento():
    assert normalize_name("João") == "JOAO"
    assert normalize_name("Estagiário") == "ESTAGIARIO"


def test_normalize_name_casa_com_e_sem_acento():
    # O ponto do B6: "João" no Procedimentos e "Joao" no Planos e Macs casam.
    assert normalize_name("joão da silva") == normalize_name("Joao Da Silva")


def test_normalize_name_tira_espacos_das_bordas():
    assert normalize_name("  José  ") == "JOSE"


def test_normalize_name_aceita_nao_string():
    assert normalize_name(123) == "123"


def test_normalizar_versao_none():
    assert normalizar_versao(None) == ""


def test_normalizar_versao_numerica():
    assert normalizar_versao(4) == "4.0"
    assert normalizar_versao(4.0) == "4.0"


def test_normalizar_versao_textual():
    assert normalizar_versao("4.0") == "4.0"
    assert normalizar_versao(" v4 ") == "v4"


def test_normalizar_versao_marcador_verificar_vira_caixa_alta():
    assert normalizar_versao("verificar") == "VERIFICAR"
    assert normalizar_versao("  Verificar ") == "VERIFICAR"
    assert normalizar_versao("VERIFICAR") == "VERIFICAR"


def test_normalizar_versao_inteiro_texto_nao_ganha_ponto_zero():
    # int/float viram "16.0"; a MESMA versão como texto fica "16".
    assert normalizar_versao(16) == "16.0"
    assert normalizar_versao("16") == "16"


def test_normalizar_versao_virgula_decimal_nao_e_canonizada():
    # DOC/BUG LATENTE: "9,0" (vírgula decimal BR) digitado como TEXTO na coluna F
    # não é convertido para "9.0" e não passa em versao_e_numerica -> o sync
    # trataria essa linha como "versão inválida" e a ignoraria (+ warning).
    # Números vindos de célula numérica do Excel chegam como float e não caem
    # aqui; só pega quem digita a versão como texto com vírgula.
    assert normalizar_versao("9,0") == "9,0"
    assert versao_e_numerica(normalizar_versao("9,0")) is False


def test_parse_data_none():
    assert parse_data(None) is None


def test_parse_data_datetime_passa_direto():
    valor = datetime(2026, 8, 10)
    assert parse_data(valor) is valor


def test_parse_data_texto_dd_mm_aaaa():
    assert parse_data("10/08/2026") == datetime(2026, 8, 10)


def test_parse_data_texto_com_aspa_de_protecao():
    assert parse_data("'10/08/2026") == datetime(2026, 8, 10)


def test_parse_data_texto_iso():
    assert parse_data("2026-08-10") == datetime(2026, 8, 10)


def test_parse_data_prioriza_leitura_dia_mes_ano():
    # 03/05 é ambíguo: deve ser lido como 3 de maio (padrão BR), não 5 de março.
    assert parse_data("03/05/2026") == datetime(2026, 5, 3)


def test_parse_data_texto_invalido():
    assert parse_data("não é uma data") is None


def test_formatar_data_texto_a_partir_de_datetime():
    assert formatar_data_texto(datetime(2026, 8, 5)) == "'05/08/2026"


def test_formatar_data_texto_a_partir_de_texto():
    assert formatar_data_texto("10/08/2026") == "'10/08/2026"


def test_formatar_data_texto_none_escreve_o_literal_none():
    # DOC/BUG LATENTE: sem data, a função devolve o texto "'None" e isso vai
    # parar na célula "planejado". Hoje só é chamada com uma data de verdade
    # (linha ON TIME tem prazo), mas se algum dia receber None a planilha ganha
    # a string "None" numa célula de data.
    assert formatar_data_texto(None) == "'None"


def test_versao_e_numerica():
    assert versao_e_numerica("9.0") is True
    assert versao_e_numerica("16") is True
    assert versao_e_numerica(normalizar_versao(8.0)) is True
    assert versao_e_numerica("VERIFICAR") is False
    assert versao_e_numerica("N/A") is False
    assert versao_e_numerica("") is False
