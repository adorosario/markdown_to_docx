# Use the official Python image from Docker Hub
FROM python:3.10.5

# Set the working directory in the container
WORKDIR /app

# Copy the current directory contents into the container at /app
COPY . /app

# Install any needed packages
RUN pip install --no-cache-dir markdown beautifulsoup4 python-docx requests

# Command to run the Python script
CMD ["python", "index.py"]
