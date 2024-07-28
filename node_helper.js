/*
  Node Helper module for MMM-MicrosoftToDo

  Purpose: Microsoft's OAutht 2.0 Token API endpoint does not support CORS,
  therefore we cannot make AJAX calls from the browser without disabling
  webSecurity in Electron.
*/
var NodeHelper = require("node_helper");
const Log = require("logger");
const Client = require("./MicrosoftToDoClient");

const clients = {};
const intervals = {};

module.exports = NodeHelper.create({
  init: function () {
    Log.info(`[MMM-MicrosoftToDo] node_helper init ...`);
  },

  start: function () {
    Log.info(`[${this.name}] node_helper started ...`);
  },

  stop: function () {
    Object.keys(intervals).forEach(clearInterval);
    Log.info(`[${this.name}] node_helper shutting down ...`);
  },

  socketNotificationReceived: function (notification, payload) {
    Log.debug(`server --> ${notification} ${payload.id}`);
    if (notification === "FETCH_DATA") {
      if (!intervals[payload.id]) {
        var self = this;
        this.fetchData(payload);
        intervals[payload.id] = setInterval(() => self.fetchData(payload), payload.refreshSeconds * 1000);
        Log.debug(`[${payload.id}] interval started with the id ${intervals[payload.id]}`);
      } else {
        Log.debug(`[${payload.id}] interval exists with the id ${intervals[payload.id]}`);
      }
    } else if (notification === "COMPLETE_TASK") {
      this.completeTask(payload.listId, payload.taskId, payload.config);
    } else {
      Log.warn(`[${config.id}] - did not process event: ${notification}`);
    }
  },

  getClient: function (config) {
    if (!clients[config.id]) {
      clients[config.id] = new Client(config);
    }
    return clients[config.id];
  },

  completeTask: function (listId, taskId, config) {
    Log.error(`[${config.id}] completeTask are not implemented yet`);
  },

  fetchData: function (config) {
    const self = this;
    const client = this.getClient(config);
    client
      .getTodos(config)
      .then(tasks => {
        self.sendSocketNotification(`DATA_FETCHED_${config.id}`, tasks);
      })
      .catch(error => {
        Log.error(`[${config.id}] - ${error}`);
        self.sendSocketNotification(`FETCH_INFO_ERROR_${config.id}`, error);
      });
  },
});
