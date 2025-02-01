input = {
    "row": "bg-white dark:bg-gray-800",
    "col": "py-2 px-4"
}

output = {
    "row": "bg-gray-100 dark:bg-gray-600",
    "col": "py-2 px-4"
}

styles = output
allow_input_on_output = True

formulas = "export var formulas = {formulas};"
data = "export var data = {data};"
row_template = f'<tr class="{styles["row"]}">{{columns}}</tr>'
col_template = f'<td class="{styles["col"]}">{{text}}</td>'
output_template = f'<td class="{styles["col"]}"><span id="{{cell_id}}">{{text}}</span></td>'
#output_template = f'<td class="{styles["col"]}">${{{cell_id}}}</td>'
input_template = f'<td class="{styles["col"]}"><input type="number" step="any" name="{{cell_id}}" class="xlsx-calc form-control w-full bg-gray-50 dark:bg-gray-700 text-gray-800 dark:text-gray-200 p-2 px-1 rounded-md" id="{{cell_id}}" min="0" value="{{value}}" required></td>'

html_template = """
<!doctype html>
<html lang="es">
  <head>
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous">
    <title>Converted</title>
    <style>
    body{{
      font-size: 1em;
    }};
    </style>
  </head>
  <body>
    <div class="container-fluid">
    <table>
      {body}
    </table>
    </div>
    <!-- Optional JavaScript; choose one of the two! -->
    <!-- Option 1: Bootstrap Bundle with Popper -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
    <!-- Option 2: Separate Popper and Bootstrap JS -->
    <!--
    <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.9.2/dist/umd/popper.min.js" integrity="sha384-IQsoLXl5PILFhosVNubq5LC7Qb9DXgDA9i+tQ8Zj3iwWAwPtgFTxbJ8NT4GN1R8p" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.min.js" integrity="sha384-cVKIPhGWiC2Al4u+LWgxfKTRIcfu0JTxR+EQDz/bgldoEyl4H0zUF0QKbrJ0EcQF" crossorigin="anonymous"></script>
    -->
  </body>
</html>
"""