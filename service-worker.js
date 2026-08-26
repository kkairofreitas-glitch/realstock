const CACHE_NAME = 'realstock-mobile-v10';

const URLS_MOBILE_CACHE = [
  '/contagem-mobile',
  '/contagem-mobile.html',
  '/manifest.json',
  '/logo-realstock.png',
  '/icones/icone-192.png',
  '/icones/icone-512.png',
];

/*
  Instala somente os arquivos do mobile.

  Não coloque:
  "/"
  "/index.html"
  "/login"
*/
self.addEventListener(
  'install',
  (event) => {
    event.waitUntil(
      caches
        .open(CACHE_NAME)
        .then(async (cache) => {
          for (
            const url
            of URLS_MOBILE_CACHE
          ) {
            try {
              await cache.add(url);
            } catch (erro) {
              console.warn(
                'Não foi possível armazenar no cache:',
                url
              );
            }
          }
        })
    );

    self.skipWaiting();
  }
);

/*
  Apaga os caches antigos, incluindo
  realstock-mobile-v8, que armazenou
  indevidamente a rota "/".
*/
self.addEventListener(
  'activate',
  (event) => {
    event.waitUntil(
      caches
        .keys()
        .then((chaves) =>
          Promise.all(
            chaves
              .filter(
                (chave) =>
                  chave !== CACHE_NAME
              )
              .map(
                (chave) =>
                  caches.delete(chave)
              )
          )
        )
        .then(() =>
          self.clients.claim()
        )
    );
  }
);

/*
  Intercepta exclusivamente a área mobile.

  O painel administrativo, login, APIs,
  exportações e demais páginas seguem
  diretamente para o servidor.
*/
self.addEventListener(
  'fetch',
  (event) => {
    const request =
      event.request;

    if (
      request.method !== 'GET'
    ) {
      return;
    }

    const url =
      new URL(request.url);

    const pertenceAoMobile =
      url.origin ===
        self.location.origin &&
      (
        url.pathname ===
          '/coleta-mobile' ||
        url.pathname ===
          '/contagem-mobile.html' ||
        url.pathname ===
          '/manifest.json' ||
        url.pathname ===
          '/logo-realstock.png' ||
        url.pathname.startsWith(
          '/icones/'
        )
      );

    /*
      Não use respondWith fora do mobile.
      Assim "/" e "/login" não são
      controlados pelo Service Worker.
    */
    if (!pertenceAoMobile) {
      return;
    }

    /*
      Para a página mobile:
      tenta a versão atual do servidor
      e usa o cache somente se a rede falhar.
    */
    event.respondWith(
      fetch(request)
        .then((resposta) => {
          if (
            resposta &&
            resposta.ok
          ) {
            const copia =
              resposta.clone();

            caches
              .open(CACHE_NAME)
              .then((cache) =>
                cache.put(
                  request,
                  copia
                )
              );
          }

          return resposta;
        })
        .catch(async () => {
          const armazenada =
            await caches.match(
              request
            );

          if (armazenada) {
            return armazenada;
          }

          if (
            url.pathname ===
            '/coleta-mobile'
          ) {
            return caches.match(
              '/contagem-mobile.html'
            );
          }

          throw new Error(
            'Recurso mobile indisponível.'
          );
        })
    );
  }
);