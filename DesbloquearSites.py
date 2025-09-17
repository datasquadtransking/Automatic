import os

# Lista dos sites que foram bloqueados
sites_bloqueados = [
    "www.facebook.com", "facebook.com",
    "www.youtube.com", "youtube.com",
    "www.instagram.com", "instagram.com"
]

# Caminho do arquivo hosts
hosts_path = r"C:\Windows\System32\drivers\etc\hosts"

def desbloquear_sites():
    try:
        with open(hosts_path, 'r+') as file:
            linhas = file.readlines()
            file.seek(0)
            for linha in linhas:
                if not any(site in linha for site in sites_bloqueados):
                    file.write(linha)
            file.truncate()
        print("Sites desbloqueados com sucesso.")
    except PermissionError:
        print("❌ Permissão negada. Execute este script como administrador.")

if __name__ == "__main__":
    desbloquear_sites()
