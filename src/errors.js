'use strict';

class ConfigError extends Error {
  constructor(message, location) {
    super(location ? `${location}: ${message}` : message);
    this.name = 'ConfigError';
    this.detail = message;
    this.location = location;
  }
}

class ConfigValidationError extends ConfigError {
  constructor(errors) {
    const normalized = errors.flatMap((error) => (
      error instanceof ConfigValidationError ? error.errors : [error]
    ));
    const message = [
      `发现 ${normalized.length} 个配置错误：`,
      ...normalized.map((error, index) => `${index + 1}. ${error.message}`)
    ].join('\n');
    super(message);
    this.name = 'ConfigValidationError';
    this.errors = normalized;
  }
}

module.exports = { ConfigError, ConfigValidationError };
