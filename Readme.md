### Exporting from ElasticSearch to CSV

Occasionally it's useful to extract the data you've just pushed into ElasticSearch. We can query ES and apply sorting, saving the resulting JSON like so:

```bash
curl -X GET 'http://localhost:9200/data_elements/data_element/_search' -H 'Content-Type: application/json' -d '
{
    "from" : 0, "size" : 100,
    "query": {
        "query_string" : {
            "query" : "*FSM*"
        }
    },
    "sort": [
      {
        "group_name.keyword": "desc"
      },
      {
        "table_name.keyword": "desc"
      },
      {
        "NPD Alias.keyword": "desc"
      }
    ]
}' > fsm_results.json
```

And then do a touch of Ruby to convert that to CSV:

```bash
ruby -rjson -rcsv -e '
  data = JSON.parse(File.read "fsm_results.json")
  data["hits"]["hits"]
    .map { |result| result["_source"] }
    .map { |result| [ result["group_name"], 
                      result["table_name"], 
                      result["NPD Alias"].join, 
                      result["Description"]
                    ] }
    .each {|row| puts row.to_csv}
' > fsm_results.csv
```