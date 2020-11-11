import Axios, { AxiosError, AxiosInstance, AxiosPromise, AxiosRequestConfig, AxiosResponse } from 'axios';
import auth, { Auth, CloudType } from './Auth';
import { Logger } from './cli';
import Utils from './Utils';
const packageJSON = require('../package.json');

class Request {
  private req: AxiosInstance;
  private _logger?: Logger;
  private _debug: boolean = false;

  public set debug(debug: boolean) {
    this._debug = debug;

    if (this._debug) {
      this.req.interceptors.request.use((config: AxiosRequestConfig): AxiosRequestConfig => {
        if (this._logger) {
          this._logger.log('Request:');
          this._logger.log(Utils.filterObject(config, ['url', 'method', 'headers', 'data', 'responseType', 'decompress']));
        }
        return config;
      });
      // since we're stubbing requests, response interceptor is never called in
      // tests, so let's exclude it from coverage
      /* c8 ignore next 13 */
      this.req.interceptors.response.use((response: AxiosResponse): AxiosResponse => {
        if (this._logger) {
          this._logger.log('Response:');
          this._logger.log(Utils.filterObject(response, ['data', 'status', 'statusText', 'headers']));
        }
        return response;
      }, (error: AxiosError): void => {
        if (this._logger) {
          this._logger.log('Request error');
          this._logger.log(Utils.filterObject(error.response, ['data', 'status', 'statusText', 'headers']));
        }
        throw error;
      });
    }
  }

  public set logger(logger: Logger) {
    this._logger = logger;
  }

  constructor() {
    this.req = Axios.create({
      headers: {
        'user-agent': `NONISV|SharePointPnP|Office365CLI/${packageJSON.version}`,
        'accept-encoding': 'gzip, deflate'
      },
      decompress: true,
      responseType: 'text',
      /* c8 ignore next */
      transformResponse: [data => data]
    });
    // since we're stubbing requests, request interceptor is never called in
    // tests, so let's exclude it from coverage
    /* c8 ignore next 7 */
    this.req.interceptors.request.use((config: AxiosRequestConfig): AxiosRequestConfig => {
      if (config.responseType === 'json') {
        config.transformResponse = Axios.defaults.transformResponse;
      }

      return config;
    });
  }

  public post<TResponse>(options: AxiosRequestConfig): Promise<TResponse> {
    options.method = 'POST';
    return this.execute(options);
  }

  public get<TResponse>(options: AxiosRequestConfig): Promise<TResponse> {
    options.method = 'GET';
    return this.execute(options);
  }

  public patch<TResponse>(options: AxiosRequestConfig): Promise<TResponse> {
    options.method = 'PATCH';
    return this.execute(options);
  }

  public put<TResponse>(options: AxiosRequestConfig): Promise<TResponse> {
    options.method = 'PUT';
    return this.execute(options);
  }

  public delete<TResponse>(options: AxiosRequestConfig): Promise<TResponse> {
    options.method = 'DELETE';
    return this.execute(options);
  }

  public head<TResponse>(options: AxiosRequestConfig): Promise<TResponse> {
    options.method = 'HEAD';
    return this.execute(options);
  }

  private execute<TResponse>(options: AxiosRequestConfig, resolve?: (res: TResponse) => void, reject?: (error: any) => void): Promise<TResponse> {
    if (!this._logger) {
      return Promise.reject('Logger not set on the request object');
    }

    this.updateRequestForCloudType(options, auth.service.cloudType);

    return new Promise<TResponse>((_resolve: (res: TResponse) => void, _reject: (error: any) => void): void => {
      const resource: string = Auth.getResourceFromUrl(options.url as string);

      ((): Promise<string> => {
        if (options.headers && options.headers['x-anonymous']) {
          return Promise.resolve('');
        }
        else {
          return auth.ensureAccessToken(resource, this._logger as Logger, this._debug)
        }
      })()
        .then((accessToken: string): AxiosPromise<TResponse> => {
          if (options.headers) {
            if (options.headers['x-anonymous']) {
              delete options.headers['x-anonymous'];
            }
            else {
              options.headers.authorization = `Bearer ${accessToken}`;
            }
          }
          return this.req(options);
        })
        .then((res: any): void => {
          if (resolve) {
            resolve(options.responseType === 'stream' ? res : res.data);
          }
          else {
            _resolve(options.responseType === 'stream' ? res : res.data);
          }
        }, (error: AxiosError): void => {
          if (error && error.response &&
            (error.response.status === 429 ||
              error.response.status === 503)) {
            let retryAfter: number = parseInt(error.response.headers['retry-after'] || '10');
            if (isNaN(retryAfter)) {
              retryAfter = 10;
            }
            if (this._debug) {
              (this._logger as Logger).log(`Request throttled. Waiting ${retryAfter}sec before retrying...`);
            }
            setTimeout(() => {
              this.execute(options, resolve || _resolve, reject || _reject);
            }, retryAfter * 1000);
          }
          else {
            if (reject) {
              reject(error);
            }
            else {
              _reject(error);
            }
          }
        });
    });
  }

  private updateRequestForCloudType(options: AxiosRequestConfig, cloudType: CloudType): void {
    if (!options.url) {
      return;
    }

    const resource: string = Auth.getResourceFromUrl(options.url);
    const cloudUrl: string = Auth.getEndpointForResource(resource, cloudType);
    options.url = options.url.replace(resource, cloudUrl);
  }
}

export default new Request();
