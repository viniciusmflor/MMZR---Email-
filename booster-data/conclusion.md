## 🐛 fix | medium [ID:conclusion-1749663924907]
**Why:** O email estava sendo enviado corretamente localmente, mas a logo não aparecia quando enviado pelo Outlook. Isso acontecia porque a logo estava sendo referenciada com caminho relativo (documentos/img/logo-MMZR-azul.png), que não funciona em clientes de email, pois eles não conseguem acessar arquivos locais do sistema.
**What:** Implementei a conversão automática da logo para base64 no MMZREmailGenerator. Adicionei o método _load_logo_as_base64() que carrega a imagem local e a converte para string base64, incorporando-a diretamente no HTML do email. Também adicionei fallback textual caso a logo não seja encontrada. Agora a logo é incluída como data:image/png;base64,... garantindo compatibilidade total com Outlook e outros clientes de email.
**Files:** `mmzr_email_generator.py`
<!-- metadata:conclusion-1749663924907 -->
