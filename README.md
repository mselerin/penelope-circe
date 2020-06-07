# penelope-circe

![Circe](https://github.com/mselerin/penelope-circe/raw/master/images/circe-96.jpg)

Circe est une extension Firefox pour l'aide à l'encodage des notes d'examens dans l'application Penelope.

## Build
* `npm run watch` pour lancer le build des sources avec un watch
* `npm run start` pour lancer firefox pour tester l'extension (bien faire un `watch` avant)
* `npm run build` pour lancer le build des sources + packaging

## Packaging
* `npx web-ext sign --api-key <WEB_EXT_API_KEY> --api-secret <WEB_EXT_API_SECRET>`

Les clés d'API peuvent s'obtenir avec un compte sur https://addons.mozilla.org/
