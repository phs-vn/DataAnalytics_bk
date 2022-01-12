from request import *


def listen_batch_job(
  batch_type: str,
  run_time=None,
  wait_time=15,
):
  """
  This function listen daily batch job

  :param run_time: accept datetime object, None or 'now' to run now
  :param batch_type: accept either 'mid' (for mid-day batch) or 'end' (for end-day batch)
  :param wait_time: number of seconds between retries
  """

  if run_time == 'now' or run_time is None:
    t = dt.datetime.now()
  else:
    t = run_time

  day = t.day
  month = t.month
  year = t.year

  while True:

    batch_status = pd.read_sql(
      f"""
            SELECT * 
            FROM 
                [batch]
            WHERE
                [batch].[date] = '{year}-{month}-{day}'
            """,
      connect_DWH_CoSo
    )

    if batch_type == 'mid':
      n = 1
    elif batch_type == 'end':
      n = 2
    else:
      raise ValueError("batch_type must be either 'mid' or 'end'")

    if batch_status.shape[0] >= n:
      break
    else:
      time.sleep(wait_time)
