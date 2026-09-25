export interface QueryPayload {
  model: string;
  view: string;
  fields: string[];
  filters?: Record<string, string> | null;
  sorts?: string[];
  limit?: string;
  pivots?: string[] | null;
}

export interface QueryTaskController {
  taskId: string;
  cancel: () => Promise<void>;
  waitForResults: (onStatus: (status: string, elapsedSeconds: number) => void) => Promise<any[]>;
}

/**
 * Executes a Looker query asynchronously via the Query Task API with live polling and cancellation.
 */
export async function runQueryTask(
  baseUrl: string,
  token: string,
  queryPayload: QueryPayload
): Promise<QueryTaskController> {
  const root = baseUrl.replace(/\/$/, "");
  const headers = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  };

  // 1. Create the immutable Looker Query object to obtain a query_id
  const queryRes = await fetch(`${root}/api/4.0/queries`, {
    method: "POST",
    headers,
    body: JSON.stringify(queryPayload),
  });

  if (!queryRes.ok) {
    const errText = await queryRes.text();
    throw new Error(`Failed to create query (${queryRes.status}): ${errText}`);
  }

  const queryObj = await queryRes.json();
  const queryId = String(queryObj.id);

  // 2. Launch asynchronous Query Task with query_id
  const createRes = await fetch(`${root}/api/4.0/query_tasks`, {
    method: "POST",
    headers,
    body: JSON.stringify({
      query_id: queryId,
      result_format: "json",
    }),
  });

  if (!createRes.ok) {
    const errText = await createRes.text();
    throw new Error(`Failed to create query task (${createRes.status}): ${errText}`);
  }

  const { id: taskId } = await createRes.json();
  let isCanceled = false;
  const abortController = new AbortController();

  const cancel = async () => {
    isCanceled = true;
    abortController.abort();
  };

  const waitForResults = async (
    onStatus: (status: string, elapsedSeconds: number) => void
  ): Promise<any[]> => {
    const startTime = Date.now();

    while (!isCanceled) {
      const elapsed = Math.round((Date.now() - startTime) / 1000);

      try {
        // Poll task info
        const statusRes = await fetch(`${root}/api/4.0/query_tasks/${taskId}`, {
          headers,
          signal: abortController.signal,
        });
        if (!statusRes.ok) {
          throw new Error(`Failed to check query task status (${statusRes.status})`);
        }

        const taskInfo = await statusRes.json();
        const status = taskInfo.status || "running";
        onStatus(status, elapsed);

        if (status === "complete") {
          const resultsRes = await fetch(
            `${root}/api/4.0/query_tasks/${taskId}/results`,
            { headers, signal: abortController.signal }
          );
          if (!resultsRes.ok) {
            throw new Error(`Failed to fetch query results (${resultsRes.status})`);
          }
          return await resultsRes.json();
        } else if (status === "error" || status === "failed" || status === "killed") {
          throw new Error(`Looker query failed with status: ${status}`);
        }
      } catch (err: any) {
        if (isCanceled || err?.name === "AbortError") {
          throw new Error("Query was canceled by the user.");
        }
        throw err;
      }

      // Poll interval
      await new Promise((resolve) => setTimeout(resolve, 800));
    }

    throw new Error("Query was canceled by the user.");
  };

  return { taskId, cancel, waitForResults };
}

/**
 * Fallback direct inline query runner for quick small queries.
 */
export async function runInlineQuery(
  baseUrl: string,
  token: string,
  queryPayload: QueryPayload
): Promise<any[]> {
  const url = `${baseUrl.replace(/\/$/, "")}/api/4.0/queries/run/json`;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(queryPayload),
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`Query failed (${res.status}): ${errText}`);
  }

  return await res.json();
}
