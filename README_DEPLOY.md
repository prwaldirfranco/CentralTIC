# Central TIC — Publicar online (grátis) no Render

## Aviso importante
Esta aplicação é de uso operacional interno. Na internet aberta:
- use **senhas fortes**
- não compartilhe a URL em grupos públicos
- o plano free pode **apagar dados** ao reiniciar (planilha, usuários, modelo)

Para uso real da equipe, prefira rede interna ou acesso restrito (VPN/túnel).

---

## 1. Estrutura mínima no GitHub

```
CentralTIC22/
  central_tic_v4_corrigido.py
  requirements.txt
  render.yaml
  templates/
    index.html
    admin.html
    aplicacoes.html
    dialogo.html
    ...
  static/          (se existir)
  dados/
    aplicacoes.xlsx
    usuarios.json
    modelo_dialogo.json
    scripts.json
    ...
```

Crie um repositório no GitHub e envie esses arquivos.

---

## 2. Conta no Render

1. Acesse https://render.com e entre com GitHub  
2. **New +** → **Web Service**  
3. Conecte o repositório da Central TIC  
4. Configure:

| Campo | Valor |
|--------|--------|
| Name | `central-tic` (ou outro) |
| Runtime | Python |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `python central_tic_v4_corrigido.py` |
| Instance | Free |

5. Variáveis de ambiente (Environment):

| Key | Value |
|-----|--------|
| `HOST` | `0.0.0.0` |
| `FORCE_PORT` | `1` |
| `COOKIE_SECURE` | `1` |

O Render define `PORT` sozinho — não precisa criar.

6. **Create Web Service** e aguarde o deploy.

7. Abra a URL gerada, exemplo:  
   `https://central-tic-xxxx.onrender.com`

---

## 3. Login inicial

Se `dados/usuarios.json` foi enviado no repositório, use os usuários de lá.  
Senões padrão (troque depois):

- `admin` / `admin123`
- `agente` / `agente123`

**Troque as senhas assim que entrar.**

---

## 4. Comportamento do plano free

- O app **dorme** sem acesso (primeira abertura pode demorar ~1 min)
- Arquivos em disco podem **não persistir** para sempre no free
- Uploads (Excel, modelo) podem ser perdidos após rebuild — guarde backup local

---

## 5. Rodar local (igual antes)

```bash
python central_tic_v4_corrigido.py
```

Abre em `http://localhost:8895`

---

## 6. Problemas comuns

**Site não abre / 502**  
- Veja logs no Render (Logs)  
- Confirme Start Command: `python central_tic_v4_corrigido.py`

**Login não grava sessão**  
- Confirme `COOKIE_SECURE=1` no HTTPS do Render  

**Aplicações vazias**  
- Envie `dados/aplicacoes.xlsx` no repositório ou faça upload pelo Admin depois do deploy  

**Dados sumiram**  
- Normal em reinício no free sem disco persistente — use backup da pasta `dados/` e `backup/`

---

## Alternativa mais segura (sem internet aberta)

- **Cloudflare Tunnel** ou **Tailscale** no PC da operação  
- Só quem você autorizar acessa, sem publicar o sistema para o mundo
