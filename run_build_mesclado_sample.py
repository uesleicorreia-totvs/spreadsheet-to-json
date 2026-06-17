import json
from services import build_mesclado_for_xlsx_by_cte, excel_bytes_from_records

# Payload (exemplo do usuário, reduzido para as chaves essenciais)
payload = {
  "notas": [
    {
      "num_nfe": "298343",
      "cte_origem": "123055"
    },
    {
      "num_nfe": "493009",
      "cte_origem": "1960"
    },
    {
      "num_nfe": "493010",
      "cte_origem": "1960"
    },
    {
      "num_nfe": "493011",
      "cte_origem": "1960"
    },
    {
      "num_nfe": "493012",
      "cte_origem": "1960"
    },
    {
      "num_nfe": "493013",
      "cte_origem": "1960"
    },
    {
      "num_nfe": "493014",
      "cte_origem": "1960"
    },
    {
      "num_nfe": "409289",
      "cte_origem": "409289"
    },
    {
      "num_nfe": "395470",
      "cte_origem": "119066"
    },
    {
      "num_nfe": "395471",
      "cte_origem": "119066"
    },
    {
      "num_nfe": "378300",
      "cte_origem": "1279664"
    },
    {
      "num_nfe": "387509",
      "cte_origem": "1284562"
    },
    {
      "num_nfe": "377952",
      "cte_origem": "79647629"
    },
    {
      "num_nfe": "376894",
      "cte_origem": "1056343"
    },
    {
      "num_nfe": "399393",
      "cte_origem": "1077015"
    },
    {
      "num_nfe": "399395",
      "cte_origem": "1077015"
    },
    {
      "num_nfe": "399404",
      "cte_origem": "1077015"
    },
    {
      "num_nfe": "399412",
      "cte_origem": "1077015"
    },
    {
      "num_nfe": "377226",
      "cte_origem": "370183"
    },
    {
      "num_nfe": "377227",
      "cte_origem": "370183"
    },
    {
      "num_nfe": "385253",
      "cte_origem": "372698"
    },
    {
      "num_nfe": "385254",
      "cte_origem": "372698"
    },
    {
      "num_nfe": "417066",
      "cte_origem": "12675047"
    },
    {
      "num_nfe": "417067",
      "cte_origem": "12675047"
    },
    {
      "num_nfe": "417068",
      "cte_origem": "12675047"
    },
    {
      "num_nfe": "417069",
      "cte_origem": "12675047"
    },
    {
      "num_nfe": "417070",
      "cte_origem": "12675047"
    },
    {
      "num_nfe": "417071",
      "cte_origem": "12675047"
    },
    {
      "num_nfe": "417072",
      "cte_origem": "12675047"
    },
    {
      "num_nfe": "403039",
      "cte_origem": "213013"
    },
    {
      "num_nfe": "403040",
      "cte_origem": "213013"
    },
    {
      "num_nfe": "403041",
      "cte_origem": "213013"
    },
    {
      "num_nfe": "403042",
      "cte_origem": "213013"
    },
    {
      "num_nfe": "403043",
      "cte_origem": "213013"
    },
    {
      "num_nfe": "403044",
      "cte_origem": "213013"
    },
    {
      "num_nfe": "403045",
      "cte_origem": "213013"
    },
    {
      "num_nfe": "403046",
      "cte_origem": "213013"
    },
    {
      "num_nfe": "403047",
      "cte_origem": "213013"
    },
    {
      "num_nfe": "403048",
      "cte_origem": "213013"
    },
    {
      "num_nfe": "403049",
      "cte_origem": "213013"
    },
    {
      "num_nfe": "381885",
      "cte_origem": "4711"
    },
    {
      "num_nfe": "381886",
      "cte_origem": "4711"
    },
    {
      "num_nfe": "381887",
      "cte_origem": "4711"
    },
    {
      "num_nfe": "381888",
      "cte_origem": "4711"
    },
    {
      "num_nfe": "381889",
      "cte_origem": "4711"
    },
    {
      "num_nfe": "427783",
      "cte_origem": "6952"
    },
    {
      "num_nfe": "427784",
      "cte_origem": "6952"
    },
    {
      "num_nfe": "427785",
      "cte_origem": "6952"
    },
    {
      "num_nfe": "427786",
      "cte_origem": "6952"
    },
    {
      "num_nfe": "427787",
      "cte_origem": "6952"
    },
    {
      "num_nfe": "427788",
      "cte_origem": "6952"
    },
    {
      "num_nfe": "393450",
      "cte_origem": "6584"
    },
    {
      "num_nfe": "393451",
      "cte_origem": "6584"
    },
    {
      "num_nfe": "393452",
      "cte_origem": "6584"
    },
    {
      "num_nfe": "395924",
      "cte_origem": "6584"
    },
    {
      "num_nfe": "395925",
      "cte_origem": "6584"
    },
    {
      "num_nfe": "395996",
      "cte_origem": "6584"
    },
    {
      "num_nfe": "370616",
      "cte_origem": "68893"
    },
    {
      "num_nfe": "370617",
      "cte_origem": "68893"
    },
    {
      "num_nfe": "370618",
      "cte_origem": "68893"
    },
    {
      "num_nfe": "370619",
      "cte_origem": "68893"
    },
    {
      "num_nfe": "370620",
      "cte_origem": "68893"
    },
    {
      "num_nfe": "380658",
      "cte_origem": "69493"
    },
    {
      "num_nfe": "380659",
      "cte_origem": "69493"
    },
    {
      "num_nfe": "380660",
      "cte_origem": "69493"
    },
    {
      "num_nfe": "380661",
      "cte_origem": "69493"
    },
    {
      "num_nfe": "380662",
      "cte_origem": "69493"
    },
    {
      "num_nfe": "380663",
      "cte_origem": "69493"
    },
    {
      "num_nfe": "380664",
      "cte_origem": "69493"
    },
    {
      "num_nfe": "380665",
      "cte_origem": "69493"
    },
    {
      "num_nfe": "380666",
      "cte_origem": "69493"
    },
    {
      "num_nfe": "391845",
      "cte_origem": "70193"
    },
    {
      "num_nfe": "390715",
      "cte_origem": "70453"
    },
    {
      "num_nfe": "403183",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403212",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403213",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403214",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403215",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403216",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403225",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403226",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403227",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403228",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403229",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403230",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403231",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403232",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403233",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403234",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403235",
      "cte_origem": "70706"
    },
    {
      "num_nfe": "403236",
      "cte_origem": "70706"
    }
  ],
  "itens": [
    {
      "cte_des": "1",
      "valor": "850.00",
      "cte_origem": "123055",
      "cnpj_origem": "03.594.123/0001-96",
      "cod_emissor": "53979821"
    },
    {
      "cte_des": "2",
      "valor": "360.00",
      "cte_origem": "1960",
      "cnpj_origem": "03.594.123/0001-96",
      "cod_emissor": "54119306"
    },
    {
      "cte_des": "3",
      "valor": "550.00",
      "cte_origem": "409289",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "822"
    },
    {
      "cte_des": "4",
      "valor": "300.00",
      "cte_origem": "119066",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "1652"
    },
    {
      "cte_des": "5",
      "valor": "360.00",
      "cte_origem": "1279664",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "2288"
    },
    {
      "cte_des": "6",
      "valor": "402.00",
      "cte_origem": "1284562",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "2288"
    },
    {
      "cte_des": "7",
      "valor": "55.00",
      "cte_origem": "79647629",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "2666"
    },
    {
      "cte_des": "8",
      "valor": "72.00",
      "cte_origem": "1056343",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "3266"
    },
    {
      "cte_des": "9",
      "valor": "248.00",
      "cte_origem": "1077015",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "3266"
    },
    {
      "cte_des": "10",
      "valor": "296.00",
      "cte_origem": "370183",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "3674"
    },
    {
      "cte_des": "11",
      "valor": "440.00",
      "cte_origem": "372698",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "3674"
    },
    {
      "cte_des": "12",
      "valor": "875.00",
      "cte_origem": "12675047",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "6044"
    },
    {
      "cte_des": "13",
      "valor": "1980.00",
      "cte_origem": "213013",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "62350"
    },
    {
      "cte_des": "14",
      "valor": "1100.00",
      "cte_origem": "4711",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "89623"
    },
    {
      "cte_des": "15",
      "valor": "840.00",
      "cte_origem": "6952",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "89623"
    },
    {
      "cte_des": "16",
      "valor": "1380.00",
      "cte_origem": "6584",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "53800705"
    },
    {
      "cte_des": "17",
      "valor": "1750.00",
      "cte_origem": "68893",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "53885801"
    },
    {
      "cte_des": "18",
      "valor": "2070.00",
      "cte_origem": "69493",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "53885801"
    },
    {
      "cte_des": "19",
      "valor": "580.00",
      "cte_origem": "70193",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "53885801"
    },
    {
      "cte_des": "20",
      "valor": "620.00",
      "cte_origem": "70453",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "53885801"
    },
    {
      "cte_des": "21",
      "valor": "3240.00",
      "cte_origem": "70706",
      "cnpj_origem": "03.594.123/0009-43",
      "cod_emissor": "53885801"
    }
  ]
}

records = build_mesclado_for_xlsx_by_cte(payload)
print('REGISTROS GERADOS:')
print(json.dumps(records, ensure_ascii=False, indent=2))

# Gerar bytes XLSX e salvar
xlsx_bytes = excel_bytes_from_records(records)
output_path = 'mesclado_sample.xlsx'
with open(output_path, 'wb') as f:
    f.write(xlsx_bytes)

print('\nArquivo salvo em:', output_path)
