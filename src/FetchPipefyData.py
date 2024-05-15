import requests
from Definitions import PIPEFY_API_TOKEN, PIPEFY_GRAPHQL_ENDPOINT

def fetch_pipefy_data(pipe_id, cursor=None, page_size=30):
    cursor_clause = f', after: "{cursor}"' if cursor else ""
    
    query = f"""
    query {{
      pipe(id: "{pipe_id}") {{
        phases {{
          name
          cards(first: {page_size}{cursor_clause}) {{
            edges {{
              node {{
                title
                createdAt
                fields {{
                  name
                  value
                }}
              }}
            }}
            pageInfo {{
              hasNextPage
              endCursor
            }}
          }}
        }}
      }}
    }}
    """

    headers = {'Authorization': f'Bearer {PIPEFY_API_TOKEN}'}
    response = requests.post(PIPEFY_GRAPHQL_ENDPOINT, json={'query': query}, headers=headers)
    
    if response.status_code == 200:
        data = response.json()
        if 'errors' in data:
            print(f"Erros na resposta da API: {data['errors']}")
        return data
    else:
        print(f"Erro na requisição: {response.status_code} - {response.text}")
        return None
