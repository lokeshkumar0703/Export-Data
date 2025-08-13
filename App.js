import React, { useState, useEffect } from 'react';
import { View, StyleSheet, ScrollView, Alert, ActivityIndicator } from 'react-native';
import { Button, TextInput, Text, Card, Title, IconButton, useTheme } from 'react-native-paper';
import * as FileSystem from 'expo-file-system';
import * as Sharing from 'expo-sharing';
import * as XLSX from 'xlsx';
// Hi sabari
//I am Lokesh
const FILE_NAME = 'data.xlsx';
const FILE_PATH = `${FileSystem.documentDirectory}${FILE_NAME}`;

const validateEmail = (email) => {
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(String(email).toLowerCase());
};

export default function App() {
  const [name, setName] = useState('');
  const [phone, setPhone] = useState('');
  const [email, setEmail] = useState('');
  const [details, setDetails] = useState('');
  const [city, setCity] = useState('');
  const [entries, setEntries] = useState([]);
  const [loading, setLoading] = useState(true);
  const [errors, setErrors] = useState({
    name: '',
    phone: '',
    email: '',
    city: ''
  });
  const [editingId, setEditingId] = useState(null);

  useEffect(() => {
    loadExcelData();
  }, []);

  const getCurrentDateTime = () => {
    const now = new Date();
    return {
      date: now.toLocaleDateString(),
      time: now.toLocaleTimeString(),
      timestamp: now.toISOString()
    };
  };

  const validateForm = () => {
    const newErrors = {
      name: '',
      phone: '',
      email: '',
      city: ''
    };
    let isValid = true;

    if (!name.trim()) {
      newErrors.name = 'Name is required';
      isValid = false;
    }

    if (!phone.trim()) {
      newErrors.phone = 'Phone is required';
      isValid = false;
    } else if (phone.length < 10) {
      newErrors.phone = 'Phone must be at least 10 digits';
      isValid = false;
    }

    if (!email.trim()) {
      newErrors.email = 'Email is required';
      isValid = false;
    } else if (!validateEmail(email)) {
      newErrors.email = 'Please enter a valid email';
      isValid = false;
    }

    if (!city.trim()) {
      newErrors.city = 'City is required';
      isValid = false;
    }

    setErrors(newErrors);
    return isValid;
  };

  const loadExcelData = async () => {
    try {
      const fileInfo = await FileSystem.getInfoAsync(FILE_PATH);
      
      if (fileInfo.exists) {
        const fileContent = await FileSystem.readAsStringAsync(FILE_PATH, {
          encoding: FileSystem.EncodingType.Base64,
        });
        
        const workbook = XLSX.read(fileContent, { type: 'base64' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        let jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        // Add unique IDs to entries if they don't exist
        jsonData = jsonData.map((item, index) => ({
          ...item,
          id: item.id || index.toString()
        }));
        
        setEntries(jsonData);
      } else {
        const headers = ['id', 'Name', 'Phone', 'Email', 'Details', 'City', 'Date', 'Time', 'Timestamp'];
        const ws = XLSX.utils.json_to_sheet([], { header: headers });
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Data');
        const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
        
        await FileSystem.writeAsStringAsync(FILE_PATH, excelBuffer, {
          encoding: FileSystem.EncodingType.Base64,
        });
      }
    } catch (error) {
      console.error('Error loading Excel file:', error);
      Alert.alert('Error', 'Failed to load data from Excel file');
    } finally {
      setLoading(false);
    }
  };

  const saveToExcel = async () => {
    if (!validateForm()) return;

    try {
      setLoading(true);
      const { date, time, timestamp } = getCurrentDateTime();
      
      const fileContent = await FileSystem.readAsStringAsync(FILE_PATH, {
        encoding: FileSystem.EncodingType.Base64,
      });
      
      const workbook = XLSX.read(fileContent, { type: 'base64' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      let jsonData = XLSX.utils.sheet_to_json(worksheet);
      
      if (editingId) {
        // Update existing entry
        jsonData = jsonData.map(item => 
          item.id === editingId ? {
            id: editingId,
            Name: name,
            Phone: phone,
            Email: email,
            Details: details,
            City: city,
            Date: date,
            Time: time,
            Timestamp: timestamp
          } : item
        );
      } else {
        // Add new entry
        const newEntry = {
          id: Date.now().toString(),
          Name: name,
          Phone: phone,
          Email: email,
          Details: details,
          City: city,
          Date: date,
          Time: time,
          Timestamp: timestamp
        };
        jsonData.push(newEntry);
      }
      
      const newWs = XLSX.utils.json_to_sheet(jsonData);
      const newWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWb, newWs, 'Data');
      
      const excelBuffer = XLSX.write(newWb, { bookType: 'xlsx', type: 'base64' });
      
      await FileSystem.writeAsStringAsync(FILE_PATH, excelBuffer, {
        encoding: FileSystem.EncodingType.Base64,
      });
      
      setEntries(jsonData);
      resetForm();
      
      Alert.alert('Success', editingId ? 'Data updated successfully' : 'Data saved successfully');
    } catch (error) {
      console.error('Error saving to Excel:', error);
      Alert.alert('Error', 'Failed to save data to Excel file');
    } finally {
      setLoading(false);
    }
  };

  const editEntry = (entry) => {
    setName(entry.Name);
    setPhone(entry.Phone);
    setEmail(entry.Email);
    setDetails(entry.Details || '');
    setCity(entry.City);
    setEditingId(entry.id);
  };

  const deleteEntry = async (id) => {
    try {
      setLoading(true);
      
      const fileContent = await FileSystem.readAsStringAsync(FILE_PATH, {
        encoding: FileSystem.EncodingType.Base64,
      });
      
      const workbook = XLSX.read(fileContent, { type: 'base64' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      let jsonData = XLSX.utils.sheet_to_json(worksheet);
      jsonData = jsonData.filter(item => item.id !== id);
      
      const newWs = XLSX.utils.json_to_sheet(jsonData);
      const newWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWb, newWs, 'Data');
      
      const excelBuffer = XLSX.write(newWb, { bookType: 'xlsx', type: 'base64' });
      
      await FileSystem.writeAsStringAsync(FILE_PATH, excelBuffer, {
        encoding: FileSystem.EncodingType.Base64,
      });
      
      setEntries(jsonData);
      resetForm();
      
      Alert.alert('Success', 'Entry deleted successfully');
    } catch (error) {
      console.error('Error deleting entry:', error);
      Alert.alert('Error', 'Failed to delete entry');
    } finally {
      setLoading(false);
    }
  };

  const resetForm = () => {
    setName('');
    setPhone('');
    setEmail('');
    setDetails('');
    setCity('');
    setEditingId(null);
  };

  const shareExcelFile = async () => {
    try {
      const isAvailable = await Sharing.isAvailableAsync();
      
      if (!isAvailable) {
        Alert.alert('Sharing not available on this platform');
        return;
      }
      
      await Sharing.shareAsync(FILE_PATH, {
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        dialogTitle: 'Share Excel File',
        UTI: 'com.microsoft.excel.xlsx',
      });
    } catch (error) {
      console.error('Error sharing file:', error);
      Alert.alert('Error', 'Failed to share the Excel file');
    }
  };

  if (loading) {
    return (
      <View style={styles.loaderContainer}>
        <ActivityIndicator size="large" />
        <Text>Loading...</Text>
      </View>
    );
  }

  return (
    <ScrollView style={styles.container}>
      <Card style={styles.card}>
        <Card.Content>
          <Title style={styles.title}>
            {editingId ? 'Edit Entry' : 'Add New Entry'}
          </Title>
          
          <TextInput
            label="Name *"
            value={name}
            onChangeText={setName}
            style={styles.input}
            mode="outlined"
            error={!!errors.name}
          />
          {errors.name ? <Text style={styles.errorText}>{errors.name}</Text> : null}
          
          <TextInput
            label="Phone Number *"
            value={phone}
            onChangeText={setPhone}
            style={styles.input}
            keyboardType="phone-pad"
            mode="outlined"
            error={!!errors.phone}
          />
          {errors.phone ? <Text style={styles.errorText}>{errors.phone}</Text> : null}

          <TextInput
            label="Email *"
            value={email}
            onChangeText={setEmail}
            style={styles.input}
            keyboardType="email-address"
            autoCapitalize="none"
            mode="outlined"
            error={!!errors.email}
          />
          {errors.email ? <Text style={styles.errorText}>{errors.email}</Text> : null}

          <TextInput
            label="City *"
            value={city}
            onChangeText={setCity}
            style={styles.input}
            mode="outlined"
            error={!!errors.city}
          />
          {errors.city ? <Text style={styles.errorText}>{errors.city}</Text> : null}
          
          <TextInput
            label="Details"
            value={details}
            onChangeText={setDetails}
            style={styles.input}
            multiline
            numberOfLines={3}
            mode="outlined"
          />
          
          <View style={styles.buttonRow}>
            <Button 
              mode="contained" 
              onPress={saveToExcel}
              style={styles.button}
              disabled={loading}
            >
              {editingId ? 'Update' : 'Save'}
            </Button>
            
            {editingId && (
              <Button 
                mode="outlined" 
                onPress={resetForm}
                style={styles.button}
                disabled={loading}
              >
                Cancel
              </Button>
            )}
          </View>
          
          <Button 
            mode="outlined" 
            onPress={shareExcelFile}
            style={styles.button}
            disabled={loading}
          >
            Share Excel File
          </Button>
        </Card.Content>
      </Card>

      {entries.length > 0 && (
        <Card style={styles.card}>
          <Card.Content>
            <Title style={styles.title}>Saved Entries</Title>
            {entries.map((entry) => (
              <View key={entry.id} style={styles.entry}>
                <Text style={styles.entryText}><Text style={styles.bold}>Name:</Text> {entry.Name}</Text>
                <Text style={styles.entryText}><Text style={styles.bold}>Phone:</Text> {entry.Phone}</Text>
                <Text style={styles.entryText}><Text style={styles.bold}>Email:</Text> {entry.Email}</Text>
                <Text style={styles.entryText}><Text style={styles.bold}>City:</Text> {entry.City}</Text>
                {entry.Details && (
                  <Text style={styles.entryText}><Text style={styles.bold}>Details:</Text> {entry.Details}</Text>
                )}
                <Text style={styles.entryText}><Text style={styles.bold}>Date:</Text> {entry.Date}</Text>
                <Text style={styles.entryText}><Text style={styles.bold}>Time:</Text> {entry.Time}</Text>
                
                <View style={styles.actionButtons}>
                  <IconButton
                    icon="pencil"
                    size={20}
                    onPress={() => editEntry(entry)}
                    style={styles.editButton}
                  />
                  <IconButton
                    icon="delete"
                    size={20}
                    onPress={() => deleteEntry(entry.id)}
                    style={styles.deleteButton}
                  />
                </View>
                
                <View style={styles.divider} />
              </View>
            ))}
          </Card.Content>
        </Card>
      )}
    </ScrollView>
  );
}

const styles = StyleSheet.create({
  container: {
    flex: 1,
    padding: 16,
    backgroundColor: '#f5f5f5',
  },
  loaderContainer: {
    flex: 1,
    justifyContent: 'center',
    alignItems: 'center',
  },
  card: {
    marginBottom: 16,
    borderRadius: 12,
    elevation: 4,
    shadowColor: '#000',
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.1,
    shadowRadius: 4,
    backgroundColor: '#ffffff',
  },
  title: {
    marginBottom: 16,
    fontSize: 24,
    fontWeight: 'bold',
    color: '#2196F3',
  },
  input: {
    marginBottom: 16,
    backgroundColor: '#ffffff',
    borderRadius: 8,
  },
  button: {
    marginTop: 8,
    marginBottom: 8,
    borderRadius: 8,
    paddingVertical: 8,
  },
  buttonRow: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    marginTop: 16,
  },
  entry: {
    paddingVertical: 12,
    paddingHorizontal: 8,
    backgroundColor: '#ffffff',
    borderRadius: 8,
    marginBottom: 8,
  },
  entryText: {
    marginBottom: 6,
    fontSize: 15,
    color: '#333333',
  },
  bold: {
    fontWeight: 'bold',
    color: '#2196F3',
  },
  divider: {
    height: 1,
    backgroundColor: '#e0e0e0',
    marginVertical: 12,
  },
  errorText: {
    color: '#f44336',
    marginBottom: 8,
    marginTop: -8,
    fontSize: 12,
  },
  actionButtons: {
    flexDirection: 'row',
    justifyContent: 'flex-end',
    marginTop: 12,
  },
  editButton: {
    backgroundColor: '#bbdefb',
    marginRight: 8,
    borderRadius: 20,
  },
  deleteButton: {
    backgroundColor: '#ffcdd2',
    borderRadius: 20,
  },
});