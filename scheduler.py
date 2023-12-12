import os
import sys
import win32com.client
from datetime import datetime, timedelta
import re
from dotenv import load_dotenv


class scheduled_task:
    def __init__(
        self, name: str, path: str, description: str, scheduled_time: datetime.time
    ):
        self._name = name
        self._path = path
        self._description = description
        self._scheduled_time = scheduled_time

    def __str__(self) -> str:
        return f"{self._name}:{self._description}, in {self._path} will run at {self._scheduled_time} "
    @property
    def name(self) -> str:
        return self._name
    @name.setter
    def name(self, n: str):
        # Add validation logic 
        self._name = n
        
    @property
    def path(self) -> str:
        return self._path
    
    @path.setter
    def path(self, p: str):
        self._path = p
   
    @property
    def description(self) -> str:
        return self._description

    @description.setter
    def description(self, desc: str):
        self._description = desc

    @property
    def scheduled_time(self) -> str:
        return self._scheduled_time
    
    @scheduled_time.setter
    def scheduled_time(self, time: datetime):
        self._scheduled_time = time

    def create_scheduled_task(self):
        scheduler = win32com.client.Dispatch("Schedule.Service")
        scheduler.Connect()

        rootFolder = scheduler.GetFolder("\\")

        # Create a new task
        taskDef = scheduler.NewTask(0)
        taskDef.RegistrationInfo.Description = self.description

        # Set the trigger to run daily
        trigger = taskDef.Triggers.Create(1)  # 1 corresponds to daily trigger
        trigger.StartBoundary = self.scheduled_time

        # Set the action to run a specific script
        execAction = taskDef.Actions.Create(0)  # 0 corresponds to execute a program
        execAction.Path = sys.executable  # Use the Python interpreter
        execAction.Arguments = self.path

        # Register the task
        rootFolder.RegisterTaskDefinition(
            self.name,
            taskDef,
            6,  # Logon type: 6 corresponds to 'Run only when user is logged on'
            None,  # User and password are None for the current user
            None,
            3,  # Logon type: 3 corresponds to 'Interactive'
        )


if __name__ == "__main__":
    
    load_dotenv('/.env')
    task_name = os.environ.get('TASK_NAME')
    script_path =os.environ.get('SCRIPT_PATH')
    description=os.environ.get('DESCRIPTION')
    start_time = datetime.strptime("08:00:00", "%H:%M:%S").time()
    cur_t = scheduled_task(task_name, script_path,description, start_time)
    # cur_t.create_scheduled_task()
