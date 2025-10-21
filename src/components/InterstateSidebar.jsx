import React, { useState } from 'react';
import { interstates, getAllStates } from '../data/interstates';

const InterstateSidebar = ({ selectedInterstate, onInterstateSelect, onStateFilter }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedState, setSelectedState] = useState('');
  const [sortBy, setSortBy] = useState('name');

  const allStates = getAllStates();

  const filteredInterstates = interstates
    .filter(interstate => {
      const matchesSearch = interstate.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
                           interstate.description.toLowerCase().includes(searchTerm.toLowerCase());
      const matchesState = !selectedState || interstate.states.includes(selectedState);
      return matchesSearch && matchesState;
    })
    .sort((a, b) => {
      switch (sortBy) {
        case 'name':
          return a.name.localeCompare(b.name);
        case 'length':
          return parseInt(b.length.replace(/[^\d]/g, '')) - parseInt(a.length.replace(/[^\d]/g, ''));
        case 'states':
          return a.states.length - b.states.length;
        default:
          return 0;
      }
    });

  const handleStateChange = (state) => {
    setSelectedState(state);
    onStateFilter(state);
  };

  const handleInterstateClick = (interstate) => {
    onInterstateSelect(interstate);
  };

  return (
    <div className="interstate-sidebar">
      <div className="sidebar-header">
        <h2>US Interstate Highways</h2>
        <p>Select an interstate to highlight on the map</p>
      </div>

      <div className="sidebar-controls">
        <div className="control-group">
          <label htmlFor="search">Search Interstates:</label>
          <input
            id="search"
            type="text"
            placeholder="Search by name or description..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="search-input"
          />
        </div>

        <div className="control-group">
          <label htmlFor="state-filter">Filter by State:</label>
          <select
            id="state-filter"
            value={selectedState}
            onChange={(e) => handleStateChange(e.target.value)}
            className="state-select"
          >
            <option value="">All States</option>
            {allStates.map(state => (
              <option key={state} value={state}>{state}</option>
            ))}
          </select>
        </div>

        <div className="control-group">
          <label htmlFor="sort">Sort by:</label>
          <select
            id="sort"
            value={sortBy}
            onChange={(e) => setSortBy(e.target.value)}
            className="sort-select"
          >
            <option value="name">Name</option>
            <option value="length">Length</option>
            <option value="states">Number of States</option>
          </select>
        </div>
      </div>

      <div className="interstate-list">
        <h3>Interstate Routes ({filteredInterstates.length})</h3>
        <div className="interstate-items">
          {filteredInterstates.map((interstate) => (
            <div
              key={interstate.id}
              className={`interstate-item ${selectedInterstate && selectedInterstate.id === interstate.id ? 'selected' : ''}`}
              onClick={() => handleInterstateClick(interstate)}
              style={{ borderLeftColor: interstate.color }}
            >
              <div className="interstate-header">
                <h4>{interstate.name}</h4>
                <span className="interstate-length">{interstate.length}</span>
              </div>
              <p className="interstate-description">{interstate.description}</p>
              <div className="interstate-states">
                <strong>States:</strong> {interstate.states.join(', ')}
              </div>
            </div>
          ))}
        </div>
      </div>

      {selectedInterstate && (
        <div className="selected-interstate-details">
          <h3>Selected: {selectedInterstate.name}</h3>
          <div className="detail-item">
            <strong>Description:</strong> {selectedInterstate.description}
          </div>
          <div className="detail-item">
            <strong>Length:</strong> {selectedInterstate.length}
          </div>
          <div className="detail-item">
            <strong>States Traversed:</strong> {selectedInterstate.states.length}
          </div>
          <div className="detail-item">
            <strong>State List:</strong> {selectedInterstate.states.join(', ')}
          </div>
        </div>
      )}

      <div className="sidebar-footer">
        <p>
          <strong>Total Interstates:</strong> {interstates.length}
        </p>
        <p>
          <strong>Total States:</strong> {allStates.length}
        </p>
        <p className="footer-note">
          Data represents major interstate highways in the US highway system.
        </p>
      </div>
    </div>
  );
};

export default InterstateSidebar;
