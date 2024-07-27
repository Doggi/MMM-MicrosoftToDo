const Log = require("logger");
const {add, formatISO9075, compareAsc, parseISO} = require("date-fns");
const {RateLimit} = require("async-sema");

class MicrosoftToDoClient {
  constructor(config) {
    Log.info(`[${config.id}] new MicrosoftToDoClient created with id ${config.id}`);
    this.config = config;
    this.accessTokenJson = undefined;
    this.tokenExpiryTime = undefined;
    this.listIds = undefined;
    this.toDoListUrl = undefined;
  }

  async getTodos() {
    const accessToken = await this.#getAccessToken();
    const listIds = await this.#fetchList(accessToken, this.config.user, this.config.listName);
    const url = this.#getTodoUrl(listIds);

    const response = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: accessToken,
      },
    });

    if (response.ok) {
      Log.debug(`[${this.config.id}] - fetched new tasks`);
      const responseData = await response.json();
      var tasks = [];
      var listId = listIds;
      if (responseData.value !== null && responseData.value !== undefined && responseData.value.length > 0) {
        tasks = responseData.value.map(element => {
          var parsedDate;
          if (element !== undefined && element.dueDateTime !== undefined) {
            parsedDate = parseISO(element.dueDateTime.dateTime);
          }
          return {
            id: element.id,
            title: element.title,
            dueDateTime: element.dueDateTime,
            recurrence: element.recurrence,
            listId: listId,
            parsedDate: parsedDate,
          };
        });
      }
      return tasks;
    } else {
      throw Error(`getTodos failed with status '${response.statusText}' ${JSON.stringify(response)}`);
    }
  }

  async #getAccessToken() {
    if (!this.accessTokenJson || Date.now() >= this.tokenExpiryTime) {
      Log.debug(`[${this.config.id}] - Requesting new access token`);
      const url = `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/token`;
      const form = new URLSearchParams();
      form.append("grant_type", "client_credentials");
      form.append("client_id", this.config.clientId);
      form.append("client_secret", this.config.clientSecret);
      form.append("scope", `https://graph.microsoft.com/${this.config.scope}`);

      const response = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: form,
      });

      if (response.ok) {
        const responseData = await response.json();
        this.accessTokenJson = responseData;
        this.tokenExpiryTime = Date.now() + responseData.expires_in * 1000;
        return `${this.accessTokenJson.token_type} ${this.accessTokenJson.access_token}`;
      } else {
        throw Error(`getAccessToken failed with status '${response.statusText}' ${JSON.stringify(response)}`);
      }
    } else {
      return `${this.accessTokenJson.token_type} ${this.accessTokenJson.access_token}`;
    }
  }

  async #fetchList(accessToken, user, listName) {
    if (!this.listIds) {
      let filterClause = "";
      if (listName !== undefined && listName !== "") {
        filterClause = `displayName eq '${listName}'`;
      }
      filterClause = encodeURIComponent(filterClause).replaceAll("'", "%27");

      let filter = "";
      if (filterClause !== "") {
        filter = `&$filter=${filterClause}`;
      }

      Log.debug(`[${this.config.id}] - getting list using filter '${filter}'`);

      const url = `https://graph.microsoft.com/v1.0/users/${user}/todo/lists?$top=200${filter}`;

      const response = await fetch(url, {
        method: "GET",
        headers: {
          Authorization: accessToken,
        },
      });
      if (response.ok) {
        const responseData = await response.json();
        if (responseData.value.length > 0) {
          this.listIds = responseData.value[0].id;
          return this.listIds;
        } else {
          throw Error(`list not found '${filter}'`);
        }
      } else {
        throw Error(`fetchList failed with status '${response.statusText}' ${JSON.stringify(response)}`);
      }
    } else {
      return this.listIds;
    }
  }

  #getTodoUrl(listId) {
    if (!this.toDoListUrl) {
      let orderBy =
        // sorting by subject is not supported anymore in API v1, hence falling back to created time
        (this.config.orderBy === "subject" ? "&$orderby=createdDateTime" : "") +
        (this.config.orderBy === "createdDate" ? "&$orderby=createdDateTime" : "") +
        (this.config.orderBy === "importance" ? "&$orderby=importance desc" : "") +
        (this.config.orderBy === "dueDate" ? "&$orderby=duedatetime/datetime" : "");

      let filterClause = "status ne 'completed'";
      if (this.config.plannedTasks.enable) {
        // default values from MMM-MicrosoftToDo.js are not considered as
        // the 'plannedTasks' configuration is handled by a nested object,
        // therefore setting default values here again
        if (!this.config.plannedTasks.duration) {
          this.config.plannedTasks.duration = {weeks: 2};
        }
        // need to ignore time zone, as the API expects a date time without
        // time zone
        var pastDate = formatISO9075(add(Date.now(), this.config.plannedTasks.duration));
        filterClause += ` and duedatetime/datetime lt '${pastDate}' and duedatetime/datetime ne null`;
      }

      filterClause = encodeURIComponent(filterClause).replaceAll("'", "%27");

      this.toDoListUrl = `https://graph.microsoft.com/v1.0/users/${this.config.user}/todo/lists/${listId}/tasks?$top=${this.config.itemLimit}&$filter=${filterClause}${orderBy}`;

      return this.toDoListUrl;
    } else {
      return this.toDoListUrl;
    }
  }
}

module.exports = MicrosoftToDoClient;
