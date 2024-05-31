# Use an official Python runtime as a parent image
FROM python:3.8

# Set the working directory in the container
WORKDIR /app

# Copy the Python script, Excel file, and requirements.txt into the container
COPY ping_test_script.py ip_addresses.xlsx requirements.txt /app/

# Install network tools (e.g., ping, netstat, curl)
RUN apt-get update && apt-get install -y iputils-ping net-tools curl

# Install Python packages from requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Expose any necessary ports (if applicable)
# EXPOSE <port>

# Define the command to run your script when the container starts
CMD ["python", "ping_test_script.py"]
