# ================================================================
#  FIUTO — immagine con tutte le dipendenze
#
#  L'attrito maggiore per chi usa FIUTO non e' il tool: sono regipy,
#  python-evtx, libesedb, libpff, yara-python, PyYAML e i binari libyal. Su una
#  workstation forense appena preparata installarli e' mezza giornata, e finche'
#  mancano i moduli che li usano si degradano — dichiarandolo, ma degradandosi.
#
#  Qui ci sono tutti. L'immagine e' l'unico modo per garantire che "nessun
#  match" e "nessuna voce" significhino davvero quello, e non "mancava la
#  libreria".
#
#  Uso:
#    docker run --rm -it \
#      -v /mnt/disk:/evidence:ro \
#      -v "$PWD/report":/report \
#      --user "$(id -u):$(id -g)" \
#      ghcr.io/fabriren/fiuto /evidence --all
#
#  --user serve perche' i report restino tuoi e non di root. /evidence in :ro
#  perche' il volume di prova non va scritto nemmeno per sbaglio: l'entrypoint
#  lo verifica e avvisa se non lo e'.
# ================================================================
FROM python:3.12-slim-bookworm

LABEL org.opencontainers.image.title="FIUTO" \
      org.opencontainers.image.description="Toolkit DFIR per analisi offline di dischi Windows/Linux/macOS" \
      org.opencontainers.image.source="https://github.com/fabriren/fiuto" \
      org.opencontainers.image.licenses="MIT"

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Strumenti di sistema che i moduli invocano direttamente.
#   file, binutils    -> file(1) e strings(1), usati da quasi tutti i moduli
#   sqlite3           -> letture SQLite fuori da Python
#   ntfs-3g           -> ntfscat, per leggere gli ADS e i file speciali NTFS
#   ewf-tools         -> ewfmount, per le immagini E01
#   dislocker         -> volumi BitLocker
#   cryptsetup-bin    -> volumi LUKS
#   util-linux        -> flock (append concorrente alla timeline), losetup
#   sleuthkit         -> fls/icat/mmls, utili accanto ai moduli
#   libyal runtime    -> dipendenze condivise dei binding Python libesedb/libpff
RUN apt-get update && apt-get install -y --no-install-recommends \
        bash coreutils findutils grep sed gawk procps \
        file binutils sqlite3 ntfs-3g ewf-tools util-linux sleuthkit \
        dislocker cryptsetup-bin fuse3 \
        libsnappy1v5 libbz2-1.0 liblzma5 zlib1g \
        ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Dipendenze Python. Le versioni non sono fissate a caso: un'immagine che
# cambia parser fra due build produce report diversi sullo stesso disco, e in
# ambito peritale e' un problema. Il pin si aggiorna deliberatamente.
COPY requirements.txt /tmp/requirements.txt
RUN apt-get update && apt-get install -y --no-install-recommends \
        build-essential libsnappy-dev \
    && pip install --no-cache-dir -r /tmp/requirements.txt \
    && apt-get purge -y --auto-remove build-essential libsnappy-dev \
    && rm -rf /var/lib/apt/lists/* /tmp/requirements.txt

WORKDIR /opt/fiuto
COPY fiuto.sh build.sh ./
COPY src/ ./src/
COPY docker-entrypoint.sh /usr/local/bin/fiuto-entrypoint

# Il file distribuito e' generato da src/. Verificarlo QUI, e non solo in CI,
# impedisce che un'immagine costruita a mano da un albero sporco contenga un
# fiuto.sh diverso dai suoi sorgenti.
RUN ./build.sh --check \
    && chmod +x fiuto.sh /usr/local/bin/fiuto-entrypoint \
    && ln -s /opt/fiuto/fiuto.sh /usr/local/bin/fiuto

# La directory dei report esiste gia' ed e' scrivibile da chiunque: con --user
# l'utente arbitrario passato da chi lancia non potrebbe crearla.
RUN mkdir -p /report && chmod 777 /report
WORKDIR /report

ENTRYPOINT ["/usr/local/bin/fiuto-entrypoint"]
CMD ["--help"]
