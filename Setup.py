from distutils.core import setup
import py2exe

Archivos = [
    ('Archivos/imgs/', ['Archivos/imgs/logo.png', 'Archivos/imgs/Icono.ico', 'Archivos/imgs/Registro.png']),
    ('Clibs', ['Clibs/Autocomplete.py']),
    ('', ['Registro.db'])
    ]

setup(
    options = 
    {'py2exe': {
        "packages": [], 
        'bundle_files': 1, 
        'compressed': True,
        'includes': []
        }},
    data_files = Archivos,
    windows = [{
            "script":"Registrar.py",
            "icon_resources": [(1, "Archivos/imgs/Icono.ico")],
            "dest_base":"Registrar Sumliprob"
            }],
    package_dir = {'Packages': ''},


)