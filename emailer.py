import os
import ssl
import smtplib
import asyncio
import logging
from typing import List, Optional, Dict
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


class Emailer:
    def __init__(self, from_email: Optional[str] = None, to_emails: Optional[List[str]] = None,
                 subject: Optional[str] = None, smtp_conf: Optional[Dict] = None):
        """
        Emailer encapsula template + envio SMTP.

        - Valores preferenciais são os passados no construtor; caso faltem,
          os valores serão lidos de variáveis de ambiente.
        """
        self.smtp_conf = smtp_conf or {}
        self.from_email = from_email or os.environ.get('EMAIL_FROM')
        self.to_emails = to_emails or (os.environ.get('EMAIL_TO') or '').split(',') if os.environ.get('EMAIL_TO') else []
        self.subject = subject or os.environ.get('EMAIL_SUBJECT') or 'Relatório'

        # SMTP defaults (fallback to env)
        self.host = self.smtp_conf.get('host') or os.environ.get('SMTP_HOST')
        self.port = int(self.smtp_conf.get('port') or os.environ.get('SMTP_PORT') or 587)
        self.user = self.smtp_conf.get('user') or os.environ.get('SMTP_USER')
        self.password = self.smtp_conf.get('password') or os.environ.get('SMTP_PASSWORD')
        self.use_ssl = str(self.smtp_conf.get('use_ssl', os.environ.get('SMTP_USE_SSL', 'false'))).lower() in ('1', 'true', 'yes')
        self.starttls = str(self.smtp_conf.get('starttls', os.environ.get('SMTP_STARTTLS', 'true'))).lower() in ('1', 'true', 'yes')
        # verbose logging
        env_verbose = str(os.environ.get('EMAIL_VERBOSE', 'false')).lower() in ('1', 'true', 'yes')
        self.verbose = env_verbose

        self._logger = logging.getLogger(__name__)

    def build_html(self, filename: str, total_records: int, error_items: list) -> str:
        """Gera o corpo HTML do e-mail com duas tabelas:

        - Tabela de resumo: quantos registros processados e quantos com erro
        - Tabela de detalhes: lista das notas com erro (num_nfe, cte_origem, cod_emissor, detalhe do erro)
        """
        # Resumo
        # Contar erros por CTe Origem único (não por notas)
        if error_items:
            unique_ctes = {str(e.get('cte_origem')) for e in error_items if e.get('cte_origem') not in (None, '')}
            errors_count = len(unique_ctes)
        else:
            errors_count = 0
        
        inserted = max(total_records - errors_count, 0) if total_records is not None else ''

        summary_table = f"""
        <table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;">
          <thead>
            <tr><th>Inseridos</th><th>Erros</th></tr>
          </thead>
          <tbody>
            <tr><td style="text-align:center">{inserted}</td><td style="text-align:center">{errors_count}</td></tr>
          </tbody>
        </table>
        """

        # Detalhes dos erros
        if not error_items:
            details_html = '<p>Nenhum item com erro.</p>'
        else:
            rows = []
            for e in error_items:
                # Usar diretamente o campo 'erro' que já vem estruturado de services.py
                detalhe_msg = e.get('erro') or ''

                rows.append(
                    f"<tr>"
                    f"<td>{e.get('num_nfe', '')}</td>"
                    f"<td>{e.get('cte_origem', '')}</td>"
                    f"<td>{e.get('cod_emissor', '')}</td>"
                    f"<td>{detalhe_msg}</td>"
                    f"</tr>"
                )

            details_rows = '\n'.join(rows)
            details_html = f"""
            <table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;">
              <thead>
                <tr>
                  <th>Número da Nota</th>
                  <th>CTe Origem</th>
                  <th>Código do Emissor</th>
                  <th>Detalhe do Erro</th>
                </tr>
              </thead>
              <tbody>
                {details_rows}
              </tbody>
            </table>
            """

        html = f"""
        <html>
          <body>
            <p>Prezados,</p>
            <p>Segue em anexo o relatório de execução: <b>{filename}</b> </p>
            <p><b>Resumo CTe</b></p>
            {summary_table}
            <p><b>Detalhes dos erros - Por Número de Nota</b></p>
            {details_html}
            <br>
                <table>
                <tr>
                    <td style="text-align:center;font-style:italic;">Total de Notas Afetadas:</td>
                    <td> {len(rows)} </td>
                </tr>
              </table>
              <br>
            <p>Atenciosamente,<br/>Agente de Descargas</p>
          </body>
        </html>
        """
        return html

    def build_message(self, filename: str, attachment_bytes: bytes, error_items: list, total_records: int = None) -> MIMEMultipart:
        msg = MIMEMultipart()
        msg['From'] = self.from_email
        msg['To'] = ', '.join(self.to_emails)
        msg['Subject'] = self.subject

        html = self.build_html(filename, total_records, error_items)
        msg.attach(MIMEText(html, 'html'))

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        msg.attach(part)

        return msg

    async def send(self, filename: str, attachment_bytes: bytes, error_items: list, total_records: int = None):
        """Constrói a mensagem e envia via SMTP de forma segura e não-bloqueante."""
        if not self.host or not self.from_email or not self.to_emails:
            raise ValueError('SMTP_HOST, EMAIL_FROM e EMAIL_TO devem ser configurados (construtor ou variáveis de ambiente)')

        msg = self.build_message(filename, attachment_bytes, error_items, total_records)

        ctx = ssl.create_default_context()
        # Tentativa robusta: tente o modo preferido (use_ssl), se falhar,
        # logue e tente o modo alternativo (STARTTLS ou SSL direto).
        def _send_blocking(use_ssl_flag: bool):
            srv = None
            try:
                if use_ssl_flag or int(self.port) == 465:
                    srv = smtplib.SMTP_SSL(self.host, self.port, timeout=30, context=ctx)
                else:
                    srv = smtplib.SMTP(self.host, self.port, timeout=30)

                srv.ehlo()
                if not use_ssl_flag and self.starttls:
                    srv.starttls(context=ctx)
                    srv.ehlo()
                if self.user and self.password:
                    srv.login(self.user, self.password)
                srv.sendmail(self.from_email, self.to_emails, msg.as_string())
            finally:
                if srv:
                    try:
                        srv.quit()
                    except Exception:
                        pass

        first_try_ssl = bool(self.use_ssl)
        last_exc = None

        # primeira tentativa
        try:
            await asyncio.to_thread(_send_blocking, first_try_ssl)
            return
        except Exception as e:
            last_exc = e
            if self.verbose:
                self._logger.exception(f'Primeira tentativa de envio falhou (use_ssl={first_try_ssl})')

        # tentar alternativa (flip SSL flag)
        try:
            await asyncio.to_thread(_send_blocking, not first_try_ssl)
            return
        except Exception as e:
            if self.verbose:
                self._logger.exception(f'Segunda tentativa de envio falhou (use_ssl={not first_try_ssl})')
            # ambos falharam — levantar com contexto original
            raise RuntimeError(f'Envio SMTP falhou (tentativas SSL={first_try_ssl} e SSL={not first_try_ssl}). Erro inicial: {last_exc}; erro final: {e}')
