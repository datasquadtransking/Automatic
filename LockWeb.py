import os
import sys

# Lista de sites a bloquear
sites_bloqueados = [
 "www.bumble.com", "bumble.com",
"www.brxbrasil.com", "brxbrasil.com",
"www.brasileirinhas.com.br", "brasileirinhas.com.br",
"www.cameraprive.com", "cameraprive.com",
"www.camera.hot", "camera.hot",
"www.dailymotion.com", "dailymotion.com",
"www.deezer.com", "deezer.com",
"www.discord.com", "discord.com",
"www.disneyplus.com", "disneyplus.com",
"www.crunchyroll.com", "crunchyroll.com",
"www.facebook.com", "facebook.com",
"www.fmovies.to", "fmovies.to",
"www.happn.com", "happn.com",
"www.hbomax.com", "hbomax.com",
"www.hulu.com", "hulu.com",
"www.instagram.com", "instagram.com",
"www.loupanproducoes.com", "loupanproducoes.com",
"www.match.com", "match.com",
"www.music.youtube.com", "music.youtube.com",
"www.netflix.com", "netflix.com",
"www.okcupid.com", "okcupid.com",
"www.play.hbomax.com", "play.hbomax.com",
"www.pinterest.com", "pinterest.com",
"www.pof.com", "pof.com",
"www.pornobrasileiro.tv", "pornobrasileiro.tv",
"www.pornhub.com", "pornhub.com",
"www.povbrazil.com", "povbrazil.com",
"www.primevideo.com", "primevideo.com",
"www.reddit.com", "reddit.com",
"www.seeking.com", "seeking.com",
"www.sexyhot.com.br", "sexyhot.com.br",
"www.spotify.com", "spotify.com",
"www.tinder.com", "tinder.com",
"www.tiktok.com", "tiktok.com",
"www.twitch.tv", "twitch.tv",
"www.tumblr.com", "tumblr.com",
"www.twitter.com", "twitter.com",
"www.vimeo.com", "vimeo.com",
"www.xvideos.com", "xvideos.com",
"www.xnudes.com", "xnudes.com",
"www.xnxx.com", "xnxx.com",
"www.youtube.com", "youtube.com",
"www.zoosk.com", "zoosk.com",

]

# Caminho do arquivo hosts
hosts_path = r"C:\Windows\System32\drivers\etc\hosts"

# Endereço de redirecionamento
redirect = "127.0.0.1"

def bloquear_sites():
    try:
        with open(hosts_path, 'r+') as file:
            conteudo = file.read()
            for site in sites_bloqueados:
                if site not in conteudo:
                    file.write(f"{redirect} {site}\n")
        print("Arquivos atualizado com sucesso.")
    except PermissionError:
        print("❌ Permissão negada. Execute este script como administrador.")

if __name__ == "__main__":
    bloquear_sites()

#pyinstaller bloquear_sites.spec