# Use an official Python runtime as a parent image
FROM python:3.9-slim

# Set the working directory in the container
WORKDIR /usr/src/app

# Copy the current directory contents into the container at /usr/src/app
COPY . .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Set environment variables (optional, if necessary)
ENV PYTHONUNBUFFERED=1

# Command to run your script when the container starts
CMD ["python", "updating-sheet.py"]
