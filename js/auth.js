/* Simple client-side password gate.
   NOT truly secure — the hash is visible in source.
   Sufficient to deter casual/unauthorized access. */
(function () {
  'use strict';

  function getHash(str) {
    var hash = 0;
    for (var i = 0; i < str.length; i++) {
      hash = ((hash << 5) - hash) + str.charCodeAt(i);
      hash |= 0;
    }
    return hash.toString();
  }

  var EXPECTED = getHash('oov26icf');

  if (sessionStorage.getItem('oov_auth') === EXPECTED) return;

  // Don't redirect if on the homepage (it has its own login overlay)
  var path = window.location.pathname;
  if (path.endsWith('index.html') || path.endsWith('/')) return;

  // Redirect to homepage for login
  window.location.replace('index.html');
})();
