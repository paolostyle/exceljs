import { stat as _stat, readdir, rmdir, unlink } from 'node:fs';
import { each } from '#lib/utils/under-dash.js';

const main = {
  cleanDir(path) {
    const deferred = Promise.defer();

    const remove = (file) => {
      const myDeferred = Promise.defer();
      const myHandler = (err) => {
        if (err) {
          myDeferred.reject(err);
        } else {
          myDeferred.resolve();
        }
      };
      _stat(file, (err, stat) => {
        if (err) {
          myDeferred.reject(err);
        } else if (stat.isFile()) {
          console.log(`unlink ${file}`);
          unlink(file, myHandler);
        } else if (stat.isDirectory()) {
          main
            .cleanDir(file)
            .then(() => {
              console.log(`rmdir ${file}`);
              rmdir(file, myHandler);
            })
            .catch(myHandler);
        }
      });
      return myDeferred.promise;
    };

    readdir(path, (err, files) => {
      if (err) {
        deferred.reject(err);
      } else {
        const promises = [];
        each(files, (file) => {
          promises.push(remove(`${path}/${file}`));
        });

        Promise.all(promises)
          .then(() => {
            deferred.resolve();
          })
          .catch((error) => {
            deferred.reject(error);
          });
      }
    });

    return deferred.promise;
  },

  randomName(length = 5) {
    const text = [];
    const possible =
      'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';

    for (let i = 0; i < length; i++)
      text.push(possible.charAt(Math.floor(Math.random() * possible.length)));

    return text.join('');
  },
  randomNum(d) {
    return Math.round(Math.random() * d);
  },

  fmt: {
    number(n) {
      // output large numbers with thousands separator
      const s = n.toString();
      const l = s.length;
      const a = [];
      let r = l % 3 || 3;
      let i = 0;
      while (i < l) {
        a.push(s.substr(i, r));
        i += r;
        r = 3;
      }
      return a.join(',');
    },
  },
};

export default main;
